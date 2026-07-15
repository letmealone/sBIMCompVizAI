# -*- coding: utf-8 -*-
"""
comparison_export.py
----------------------
두 IFC(전문가 vs AI)의 층/공간 매칭 결과와, 매칭된 공간별 객체(부재) 비교, 미매칭
공간의 객체 요약을 엑셀 워크북(여러 시트)으로 만든다.
"""
from collections import Counter

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill

import floorplan_core as fc
import ifc_to_excel as ite

# 차이가 나는 셀에 적용할 배경색 (연한 붉은색)
DIFF_FILL = PatternFill(start_color="FFE6E6", end_color="FFE6E6", fill_type="solid")


def _floor_space_polygons(storey_entity, tol=0.05):
    spaces_raw = fc.get_elements_for_storey(storey_entity, classes={'IfcSpace'})
    spaces = []
    for sp in spaces_raw:
        poly = fc.get_footprint_polygon(sp, tol=tol)
        if poly is None:
            continue
        spaces.append({'guid': sp.GlobalId, 'name': sp.Name or '(이름없음)', 'polygon': poly, 'entity': sp})
    return spaces


def _element_class_counts(ifc_file, space_entity):
    elements = fc.get_space_related_elements(ifc_file, space_entity)
    return Counter(e.is_a() for e in elements)


def build_floor_mapping_df(floor_mapping, is_llm=False, llm_detail=None):
    detail_by_a = {d['a_name']: d for d in (llm_detail or [])}
    rows = []
    for a_name, b_name in floor_mapping.items():
        row = {'전문가(A) 층': a_name, 'AI(B) 층': b_name or '(대응없음)'}
        if is_llm:
            d = detail_by_a.get(a_name, {})
            row['확신도'] = d.get('confidence', '')
            row['판단근거'] = d.get('reason', '')
        rows.append(row)
    return pd.DataFrame(rows)


def compute_all_space_matches(data_a, data_b, floor_mapping, spaces_dict_a, spaces_dict_b, area_thresh=2.0, centroid_thresh=1.0, status_cb=None):
    matched_pairs, unmatched_a, unmatched_b = [], [], []
    storeys_b_by_name = {s['Name']: s for s in data_b['storeys']}

    for a_storey in data_a['storeys']:
        a_name = a_storey['Name']
        b_name = floor_mapping.get(a_name)
        if status_cb:
            status_cb(f'공간 매칭 계산 중: {a_name} ...')
        
        spaces_a = spaces_dict_a.get(a_name, [])

        if not b_name or b_name not in storeys_b_by_name:
            for sp in spaces_a:
                unmatched_a.append((a_name, sp))
            continue

        b_storey = storeys_b_by_name[b_name]
        spaces_b = spaces_dict_b.get(b_name, [])
        
        a_to_b, b_to_a, offset, match_info = fc.match_spaces(
            spaces_a, spaces_b, area_thresh=area_thresh, centroid_thresh=centroid_thresh)

        by_guid_a = {s['guid']: s for s in spaces_a}
        by_guid_b = {s['guid']: s for s in spaces_b}
        for m in match_info:
            sa, sb = by_guid_a[m['a_guid']], by_guid_b[m['b_guid']]
            matched_pairs.append({
                'A_층': a_name, 'A_공간': sa['name'], 'A_면적(㎡)': round(sa['polygon'].area, 2),
                'B_층': b_name, 'B_공간': sb['name'], 'B_면적(㎡)': round(sb['polygon'].area, 2),
                '면적차(㎡)': m['area_diff_m2'], 'centroid거리(m)': m['centroid_dist_m'],
                'A_entity': sa['entity'], 'B_entity': sb['entity'],
            })
        for sp in spaces_a:
            if sp['guid'] not in a_to_b:
                unmatched_a.append((a_name, sp))
        for sp in spaces_b:
            if sp['guid'] not in b_to_a:
                unmatched_b.append((b_name, sp))

    df_space = pd.DataFrame([
        {k: v for k, v in r.items() if k not in ('A_entity', 'B_entity')} for r in matched_pairs
    ])
    return matched_pairs, unmatched_a, unmatched_b, df_space


def build_matched_object_comparison_df(ifc_a, ifc_b, matched_pairs):
    rows = []
    for m in matched_pairs:
        cnt_a = _element_class_counts(ifc_a, m['A_entity'])
        cnt_b = _element_class_counts(ifc_b, m['B_entity'])
        row = {
            'A_층': m['A_층'], 'A_공간': m['A_공간'], 'B_층': m['B_층'], 'B_공간': m['B_공간'],
            'A_면적(㎡)': m['A_면적(㎡)'], 'B_면적(㎡)': m['B_면적(㎡)'],
        }
        for c in sorted(set(cnt_a) | set(cnt_b)):
            a_n, b_n = cnt_a.get(c, 0), cnt_b.get(c, 0)
            row[f'{c}_A개수'] = a_n
            row[f'{c}_B개수'] = b_n
        rows.append(row)
    return pd.DataFrame(rows)


def build_matched_space_detail_df(data_a, data_b, matched_pairs):
    rows = []
    for m in matched_pairs:
        detail_a = fc.build_space_detail(data_a['ifc_file'], data_a['wall_classification'], m['A_entity'])
        detail_b = fc.build_space_detail(data_b['ifc_file'], data_b['wall_classification'], m['B_entity'])
        
        base_info = {
            'A_층': m['A_층'], 'A_공간': m['A_공간'],
            'B_층': m['B_층'], 'B_공간': m['B_공간']
        }
        
        keys_wall_cnt = set(detail_a['wall_simple_counts']) | set(detail_b['wall_simple_counts'])
        for k in sorted(keys_wall_cnt):
            r = dict(base_info)
            r.update({'대분류': '벽 내/외부 구분(개수)', '상세구분': k, 
                      '전문가(A)_수치': detail_a['wall_simple_counts'].get(k, 0),
                      'AI(B)_수치': detail_b['wall_simple_counts'].get(k, 0), '단위': '개'})
            rows.append(r)
            
        keys_wall_area = set(detail_a['wall_simple_area']) | set(detail_b['wall_simple_area'])
        for k in sorted(keys_wall_area):
            r = dict(base_info)
            r.update({'대분류': '벽 내/외부 구분(면적)', '상세구분': k, 
                      '전문가(A)_수치': detail_a['wall_simple_area'].get(k, 0.0),
                      'AI(B)_수치': detail_b['wall_simple_area'].get(k, 0.0), '단위': '㎡'})
            rows.append(r)
            
        keys_area = set(detail_a['area_by_class']) | set(detail_b['area_by_class'])
        for k in sorted(keys_area):
            val_a = detail_a['area_by_class'].get(k, {}).get('면적합계(㎡)')
            val_b = detail_b['area_by_class'].get(k, {}).get('면적합계(㎡)')
            r = dict(base_info)
            r.update({'대분류': '벽 이외 부재 합산 면적', '상세구분': k, 
                      '전문가(A)_수치': val_a if val_a is not None else 0.0,
                      'AI(B)_수치': val_b if val_b is not None else 0.0, '단위': '㎡'})
            rows.append(r)
            
        keys_eq = set(detail_a['equipment_counts']) | set(detail_b['equipment_counts'])
        for k in sorted(keys_eq):
            r = dict(base_info)
            r.update({'대분류': '설비 개수', '상세구분': k, 
                      '전문가(A)_수치': detail_a['equipment_counts'].get(k, 0),
                      'AI(B)_수치': detail_b['equipment_counts'].get(k, 0), '단위': '개'})
            rows.append(r)
            
    return pd.DataFrame(rows)


def build_unmatched_object_summary_df(ifc_a, unmatched_a, ifc_b, unmatched_b):
    rows = []
    for ifc_file, unmatched_list, side_label in (
        (ifc_a, unmatched_a, '전문가(A)'), (ifc_b, unmatched_b, 'AI(B)'),
    ):
        for storey_name, sp in unmatched_list:
            cnt = _element_class_counts(ifc_file, sp['entity'])
            row = {'출처': side_label, '층': storey_name, '공간': sp['name'],
                   '면적(㎡)': round(sp['polygon'].area, 2)}
            for c, n in cnt.items():
                row[f'{c}_개수'] = n
            rows.append(row)
    return pd.DataFrame(rows)


def _apply_diff_formatting(ws, df, sheet_type):
    if df.empty:
        return

    headers = list(df.columns)
    
    def _to_float(v):
        try:
            return float(v)
        except (TypeError, ValueError):
            return 0.0

    if sheet_type == '03_객체비교':
        pairs = []
        if 'A_면적(㎡)' in headers and 'B_면적(㎡)' in headers:
            pairs.append((headers.index('A_면적(㎡)') + 1, headers.index('B_면적(㎡)') + 1))
        
        classes = set([c.replace('_A개수', '') for c in headers if c.endswith('_A개수')])
        for c in classes:
            if f'{c}_A개수' in headers and f'{c}_B개수' in headers:
                pairs.append((headers.index(f'{c}_A개수') + 1, headers.index(f'{c}_B개수') + 1))
                
        for row_idx in range(2, ws.max_row + 1):
            for col_a, col_b in pairs:
                val_a = ws.cell(row=row_idx, column=col_a).value
                val_b = ws.cell(row=row_idx, column=col_b).value
                
                if abs(_to_float(val_a) - _to_float(val_b)) > 1e-4:
                    ws.cell(row=row_idx, column=col_a).fill = DIFF_FILL
                    ws.cell(row=row_idx, column=col_b).fill = DIFF_FILL

    elif sheet_type == '03b_상세비교':
        if '전문가(A)_수치' in headers and 'AI(B)_수치' in headers:
            col_a = headers.index('전문가(A)_수치') + 1
            col_b = headers.index('AI(B)_수치') + 1
            for row_idx in range(2, ws.max_row + 1):
                val_a = ws.cell(row=row_idx, column=col_a).value
                val_b = ws.cell(row=row_idx, column=col_b).value
                
                if abs(_to_float(val_a) - _to_float(val_b)) > 1e-4:
                    ws.cell(row=row_idx, column=col_a).fill = DIFF_FILL
                    ws.cell(row=row_idx, column=col_b).fill = DIFF_FILL


def build_comparison_workbook(data_a, data_b, floor_mapping, spaces_dict_a, spaces_dict_b, is_llm_floor_mapping=False,
                               llm_floor_detail=None, area_thresh=2.0, centroid_thresh=1.0,
                               status_cb=None):
    wb = Workbook()
    wb.remove(wb.active)
    used_names = set()

    if status_cb:
        status_cb('층 매칭 시트 작성 중...')
    df_floor = build_floor_mapping_df(floor_mapping, is_llm=is_llm_floor_mapping, llm_detail=llm_floor_detail)
    ws = wb.create_sheet(ite._unique_sheet_name(wb, '01_층_매칭', used_names))
    ite._write_df(ws, df_floor)

    matched_pairs, unmatched_a, unmatched_b, df_space = compute_all_space_matches(
        data_a, data_b, floor_mapping, spaces_dict_a, spaces_dict_b, area_thresh=area_thresh, centroid_thresh=centroid_thresh,
        status_cb=status_cb,
    )
    ws = wb.create_sheet(ite._unique_sheet_name(wb, '02_공간_매칭', used_names))
    ite._write_df(ws, df_space)

    if status_cb:
        status_cb('매칭된 공간별 객체 비교 계산 중...')
    df_compare = build_matched_object_comparison_df(data_a['ifc_file'], data_b['ifc_file'], matched_pairs)
    ws_compare = wb.create_sheet(ite._unique_sheet_name(wb, '03_매칭공간_객체비교', used_names))
    ite._write_df(ws_compare, df_compare)
    _apply_diff_formatting(ws_compare, df_compare, '03_객체비교')

    if status_cb:
        status_cb('매칭된 공간별 상세 비교(면적/설비 등) 계산 중...')
    df_detail = build_matched_space_detail_df(data_a, data_b, matched_pairs)
    ws_detail = wb.create_sheet(ite._unique_sheet_name(wb, '03b_매칭공간_상세비교', used_names))
    ite._write_df(ws_detail, df_detail)
    _apply_diff_formatting(ws_detail, df_detail, '03b_상세비교')

    if status_cb:
        status_cb('미매칭 공간 객체 요약 계산 중...')
    df_unmatched = build_unmatched_object_summary_df(data_a['ifc_file'], unmatched_a, data_b['ifc_file'], unmatched_b)
    ws = wb.create_sheet(ite._unique_sheet_name(wb, '04_미매칭공간_객체요약', used_names))
    ite._write_df(ws, df_unmatched)

    return wb