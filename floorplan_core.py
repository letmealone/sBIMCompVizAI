# -*- coding: utf-8 -*-
"""
floorplan_core.py
------------------
Streamlit 평면도 비교 앱의 데이터/지오메트리 처리 로직.
ifc_to_excel.py의 함수(내외벽 판정, 면적 계산 등)를 최대한 재사용한다.
Streamlit에 의존하지 않으므로 단독으로 import/테스트 가능하다.
"""
import statistics
from collections import Counter, defaultdict

import numpy as np
import ifcopenshell
import ifcopenshell.geom as geom
from shapely.geometry import Polygon, Point
from shapely.ops import unary_union

import ifc_to_excel as ite

_SETTINGS = geom.settings()
_SETTINGS.set('use-world-coords', True)

PLAN_STRUCTURAL_CLASSES = (
    'IfcWall', 'IfcWallStandardCase', 'IfcColumn', 'IfcBeam',
    'IfcSlab', 'IfcCurtainWall', 'IfcDoor', 'IfcWindow',
)


def load_ifc(path):
    ifc_file = ifcopenshell.open(path)
    storeys = []
    for s in ifc_file.by_type('IfcBuildingStorey'):
        storeys.append({'Name': s.Name, 'Elevation': s.Elevation, 'entity': s})
    storeys.sort(key=lambda x: (x['Elevation'] is None, x['Elevation']))
    wall_classification = ite._determine_wall_classification(ifc_file)
    return {
        'ifc_file': ifc_file,
        'storeys': storeys,
        'wall_classification': wall_classification,
    }


def match_storeys(storeys_a, storeys_b, gap_cost=1000.0):
    valid_a = [s for s in storeys_a if s['Elevation'] is not None]
    valid_b = [s for s in storeys_b if s['Elevation'] is not None]
    if not valid_a or not valid_b:
        return {}, 0.0

    naive_diffs = []
    for a in valid_a:
        nearest = min(valid_b, key=lambda x: abs(x['Elevation'] - a['Elevation']))
        naive_diffs.append(a['Elevation'] - nearest['Elevation'])
    offset = statistics.median(naive_diffs)

    seq_a = [a['Elevation'] - offset for a in valid_a]
    seq_b = [b['Elevation'] for b in valid_b]
    n, m = len(seq_a), len(seq_b)

    dp = [[0.0] * (m + 1) for _ in range(n + 1)]
    for i in range(1, n + 1):
        dp[i][0] = dp[i - 1][0] + gap_cost
    for j in range(1, m + 1):
        dp[0][j] = dp[0][j - 1] + gap_cost
    for i in range(1, n + 1):
        for j in range(1, m + 1):
            match_cost = abs(seq_a[i - 1] - seq_b[j - 1])
            dp[i][j] = min(
                dp[i - 1][j - 1] + match_cost,
                dp[i - 1][j] + gap_cost,
                dp[i][j - 1] + gap_cost,
            )

    i, j = n, m
    pairs = []
    while i > 0 or j > 0:
        if i > 0 and j > 0 and abs(dp[i][j] - (dp[i - 1][j - 1] + abs(seq_a[i - 1] - seq_b[j - 1]))) < 1e-6:
            pairs.append((i - 1, j - 1)); i -= 1; j -= 1
        elif i > 0 and abs(dp[i][j] - (dp[i - 1][j] + gap_cost)) < 1e-6:
            pairs.append((i - 1, None)); i -= 1
        else:
            pairs.append((None, j - 1)); j -= 1
    pairs.reverse()

    mapping = {}
    for a_idx, b_idx in pairs:
        if a_idx is None:
            continue
        a_name = valid_a[a_idx]['Name']
        mapping[a_name] = valid_b[b_idx]['Name'] if b_idx is not None else None
    return mapping, offset


def _offset_candidates(spaces_a, spaces_b, area_thresh):
    candidates = []
    for a in spaces_a:
        ca = a['polygon'].centroid
        aa = a['polygon'].area
        for b in spaces_b:
            ab = b['polygon'].area
            if abs(aa - ab) <= area_thresh:
                cb = b['polygon'].centroid
                candidates.append((cb.x - ca.x, cb.y - ca.y))
    return candidates


def _estimate_offset(candidates, cluster_tol=0.5):
    if not candidates:
        return None
    arr = np.array(candidates)
    best_center, best_count = None, 0
    for c in arr:
        dist = np.linalg.norm(arr - c, axis=1)
        inliers = arr[dist <= cluster_tol]
        if len(inliers) > best_count:
            best_count = len(inliers)
            best_center = inliers.mean(axis=0)
    return (float(best_center[0]), float(best_center[1])) if best_center is not None else None


def match_spaces(spaces_a, spaces_b, area_thresh=2.0, centroid_thresh=1.0):
    candidates = _offset_candidates(spaces_a, spaces_b, area_thresh)
    offset = _estimate_offset(candidates)
    if offset is None:
        return {}, {}, None, []
    dx, dy = offset

    pairs = []
    for a in spaces_a:
        ca = a['polygon'].centroid
        aa = a['polygon'].area
        for b in spaces_b:
            cb = b['polygon'].centroid
            ab = b['polygon'].area
            area_diff = abs(aa - ab)
            if area_diff > area_thresh:
                continue
            dist = ((ca.x + dx - cb.x) ** 2 + (ca.y + dy - cb.y) ** 2) ** 0.5
            if dist <= centroid_thresh:
                pairs.append((dist, a['guid'], b['guid'], area_diff))

    pairs.sort(key=lambda p: p[0])
    used_a, used_b = set(), set()
    a_to_b, b_to_a, match_info = {}, {}, []
    for dist, ga, gb, adiff in pairs:
        if ga in used_a or gb in used_b:
            continue
        used_a.add(ga); used_b.add(gb)
        a_to_b[ga] = gb
        b_to_a[gb] = ga
        match_info.append({'a_guid': ga, 'b_guid': gb, 'centroid_dist_m': round(dist, 3), 'area_diff_m2': round(adiff, 3)})

    return a_to_b, b_to_a, offset, match_info


def get_footprint_polygon(ent, tol=0.05):
    try:
        shape = geom.create_shape(_SETTINGS, ent)
    except Exception:
        return None
    verts = np.array(shape.geometry.verts).reshape(-1, 3)
    faces = np.array(shape.geometry.faces).reshape(-1, 3)
    if len(verts) == 0 or len(faces) == 0:
        return None
    zmin = verts[:, 2].min()
    polys = []
    for tri in faces:
        p = verts[tri]
        if np.all(p[:, 2] <= zmin + tol):
            try:
                poly = Polygon(p[:, :2])
                if poly.is_valid and poly.area > 1e-9:
                    polys.append(poly)
            except Exception:
                continue
    if not polys:
        return None
    try:
        u = unary_union(polys)
    except Exception:
        return None
    if u.is_empty:
        return None
    return u


def _grid_points_in_polygon(poly, spacing=0.5):
    minx, miny, maxx, maxy = poly.bounds
    xs = np.arange(minx, maxx + spacing, spacing)
    ys = np.arange(miny, maxy + spacing, spacing)
    pts = []
    for x in xs:
        for y in ys:
            if poly.contains(Point(x, y)):
                pts.append((float(x), float(y)))
    if not pts:
        c = poly.centroid
        pts.append((float(c.x), float(c.y)))
    return pts


def _polygon_xy_lists(poly):
    xs, ys = [], []
    def _add_ring(coords):
        cx, cy = zip(*coords)
        xs.extend(cx); xs.append(None)
        ys.extend(cy); ys.append(None)

    geoms = poly.geoms if poly.geom_type == 'MultiPolygon' else [poly]
    for g in geoms:
        _add_ring(list(g.exterior.coords))
    return xs, ys


def get_elements_for_storey(storey_entity, classes=None):
    storey_ifc = storey_entity['entity'] if isinstance(storey_entity, dict) else storey_entity
    elements = []
    for rel in (storey_ifc.ContainsElements or []):
        for el in rel.RelatedElements:
            if classes is None or el.is_a() in classes:
                elements.append(el)
    for rel in (storey_ifc.IsDecomposedBy or []):
        for el in rel.RelatedObjects:
            if classes is None or el.is_a() in classes:
                elements.append(el)
    return elements


def get_storey_space_summary(storey_entity, max_names=20):
    spaces = get_elements_for_storey(storey_entity, classes={'IfcSpace'})
    names = [(sp.LongName or sp.Name or '(이름없음)') for sp in spaces]
    return {'count': len(names), 'names': names[:max_names]}


def get_element_hover_info(ent, wall_classification=None):
    lines = [f"{ent.is_a()} {ent.Name or '(이름없음)'}".strip()]
    try:
        dims = ite._dimension_columns(ent)
        dim_parts = [f"{axis}={dims[f'치수_{axis}(m)']}m"
                     for axis in ('X', 'Y', 'Z') if dims.get(f'치수_{axis}(m)') is not None]
        if dim_parts:
            lines.append('치수: ' + ', '.join(dim_parts) + f" ({dims.get('치수산출방식')})")
    except Exception:
        pass
    try:
        if ent.is_a() in ite.AREA_TARGET_CLASSES:
            flat = ite._flatten_psets(ent)
            area_info = ite._area_columns(ent, flat)
            if area_info.get('면적(㎡)') is not None:
                lines.append(f"면적: {area_info['면적(㎡)']}㎡ ({area_info.get('면적산출방식')})")
    except Exception:
        pass
    try:
        material = ite._get_material(ent)
        if material:
            lines.append(f"재질: {material}")
    except Exception:
        pass
    if wall_classification and ent.GlobalId in wall_classification:
        label, reason = wall_classification[ent.GlobalId]
        lines.append(f"내/외벽 판정: {label}")
    return "<br>".join(lines)


def build_storey_plan_data(storey_entity, tol=0.05, wall_classification=None):
    spaces_raw = get_elements_for_storey(storey_entity, classes={'IfcSpace'})
    structural_raw = get_elements_for_storey(storey_entity, classes=set(PLAN_STRUCTURAL_CLASSES))

    spaces = []
    for sp in spaces_raw:
        poly = get_footprint_polygon(sp, tol=tol)
        if poly is None:
            continue
        spaces.append({'guid': sp.GlobalId, 'name': sp.Name or '(이름없음)', 'polygon': poly, 'entity': sp})

    structural = []
    for el in structural_raw:
        poly = get_footprint_polygon(el, tol=tol)
        if poly is None:
            continue
        hover = get_element_hover_info(el, wall_classification=wall_classification)
        structural.append({'guid': el.GlobalId, 'class': el.is_a(), 'name': el.Name or '',
                            'polygon': poly, 'hover': hover})

    return {'spaces': spaces, 'structural': structural}


EQUIPMENT_CLASSES = ('IfcLightFixture', 'IfcSensor', 'IfcFireSuppressionTerminal', 'IfcAlarm')


def get_space_related_elements(ifc_file, space_entity):
    by_guid = {}
    for rel in ifc_file.by_type('IfcRelSpaceBoundary'):
        if rel.RelatingSpace == space_entity and rel.RelatedBuildingElement is not None:
            el = rel.RelatedBuildingElement
            by_guid[el.GlobalId] = el
    return list(by_guid.values())


def get_space_contained_equipment(ifc_file, space_entity, classes=EQUIPMENT_CLASSES):
    equipment = []
    for rel in ifc_file.by_type('IfcRelContainedInSpatialStructure'):
        if rel.RelatingStructure == space_entity:
            for el in rel.RelatedElements:
                if el.is_a() in classes:
                    equipment.append(el)
    return equipment


def _wall_display_category(result):
    if result == '내벽':
        return '내부', '내벽'
    if result == '내벽(추정-관계기반)':
        return '내부(추정)', result
    if result == '판정불가':
        return '외부(판정불가)', '외벽(판정불가)'
    return '외부(판정됨)', result


_AREA_APPORTION_CLASSES = ('IfcSlab', 'IfcRoof', 'IfcCovering')
_APPORTION_BUFFER_M = 0.3  


def _space_portion_fraction(member_entity, space_footprint, buffer_dist=_APPORTION_BUFFER_M):
    if space_footprint is None or space_footprint.is_empty:
        return None
    member_footprint = get_footprint_polygon(member_entity)
    if member_footprint is None or member_footprint.is_empty or member_footprint.area <= 0:
        return None
    try:
        buffered_space = space_footprint.buffer(buffer_dist)
        inter = member_footprint.intersection(buffered_space)
    except Exception:
        return None
    if inter.is_empty:
        return 0.0
    return min(inter.area / member_footprint.area, 1.0)


def build_space_detail(ifc_file, wall_classification, space_entity):
    related = get_space_related_elements(ifc_file, space_entity)
    equipment = get_space_contained_equipment(ifc_file, space_entity)
    space_footprint = get_footprint_polygon(space_entity)

    class_counts = Counter(e.is_a() for e in related)

    wall_simple_counts = Counter()
    wall_simple_area = Counter()
    wall_detail_counts = Counter()
    wall_area_apportioned = True
    for e in related:
        if not e.is_a('IfcWall'):
            continue
        result, _reason = wall_classification.get(e.GlobalId, ('판정불가', ''))
        simple, detail_label = _wall_display_category(result)
        wall_simple_counts[simple] += 1
        wall_detail_counts[detail_label] += 1
        flat = ite._flatten_psets(e)
        v = flat.get('Qto_WallBaseQuantities.Gross_Side_Area')
        if isinstance(v, (int, float)):
            fraction = _space_portion_fraction(e, space_footprint)
            if fraction is None:
                wall_area_apportioned = False
                fraction = 1.0
            wall_simple_area[simple] += v * fraction

    area_by_class = {}
    non_wall_classes = sorted(set(e.is_a() for e in related if not e.is_a('IfcWall')))
    for cls in non_wall_classes:
        ents = [e for e in related if e.is_a(cls)]
        total, n_ok = 0.0, 0
        apportioned = cls in _AREA_APPORTION_CLASSES
        any_fallback = False
        for e in ents:
            flat = ite._flatten_psets(e)
            cols = ite._area_columns(e, flat)
            if cols['면적(㎡)'] is not None:
                fraction = 1.0
                if apportioned:
                    fraction = _space_portion_fraction(e, space_footprint)
                    if fraction is None:
                        any_fallback = True
                        fraction = 1.0
                total += cols['면적(㎡)'] * fraction
                n_ok += 1
        note = ''
        if apportioned:
            note = '(공간 귀속분 안분)' if not any_fallback else '(일부 안분실패-전체값 폴백)'
        area_by_class[cls] = {
            '면적합계(㎡)': round(total, 2) if n_ok else None,
            '산출가능/전체': f'{n_ok}/{len(ents)}',
            '비고': note,
        }

    equipment_counts = Counter(e.is_a() for e in equipment)

    flat_sp = ite._flatten_psets(space_entity)
    space_area, space_area_method = None, None
    for key in ('Qto_SpaceBaseQuantities.NetFloorArea', 'Qto_SpaceBaseQuantities.GrossFloorArea'):
        if isinstance(flat_sp.get(key), (int, float)):
            space_area, space_area_method = flat_sp[key], key
            break
    if space_area is None:
        cols = ite._area_columns(space_entity, flat_sp)
        space_area, space_area_method = cols['면적(㎡)'], cols['면적산출방식']

    return {
        'name': space_entity.Name or '(이름없음)',
        'long_name': space_entity.LongName,
        'guid': space_entity.GlobalId,
        'area': round(space_area, 2) if space_area is not None else None,
        'area_method': space_area_method,
        'class_counts': dict(class_counts),
        'wall_simple_counts': dict(wall_simple_counts),
        'wall_simple_area': {k: round(v, 2) for k, v in wall_simple_area.items()},
        'wall_area_note': '(공간 귀속분 안분)' if wall_area_apportioned else '(일부 안분실패-전체값 폴백 포함)',
        'wall_detail_counts': dict(wall_detail_counts),
        'area_by_class': area_by_class,
        'equipment_counts': dict(equipment_counts),
        'highlight_map': _build_highlight_map(related, equipment, wall_classification),
    }


def _build_highlight_map(related, equipment, wall_classification):
    hl = {}
    for e in related:
        if e.is_a('IfcWall'):
            result, _ = wall_classification.get(e.GlobalId, ('판정불가', ''))
            simple, _ = _wall_display_category(result)
            if simple == '내부':
                hl[e.GlobalId] = 'wall_internal'
            elif simple == '내부(추정)':
                hl[e.GlobalId] = 'wall_internal_estimated'
            elif simple == '외부(판정불가)':
                hl[e.GlobalId] = 'wall_external_unknown'
            else:
                hl[e.GlobalId] = 'wall_external_confirmed'
        else:
            hl[e.GlobalId] = 'related'
    for e in equipment:
        hl[e.GlobalId] = 'equipment'
    return hl


_STRUCT_COLORS = {
    'IfcWall': 'rgba(90,90,90,0.85)',
    'IfcWallStandardCase': 'rgba(90,90,90,0.85)',
    'IfcColumn': 'rgba(40,40,40,0.9)',
    'IfcBeam': 'rgba(120,90,60,0.7)',
    'IfcSlab': 'rgba(200,190,170,0.4)',
    'IfcCurtainWall': 'rgba(120,170,220,0.6)',
    'IfcDoor': 'rgba(150,100,50,0.6)',
    'IfcWindow': 'rgba(120,200,230,0.6)',
}
_FADED_COLOR = 'rgba(210,210,210,0.35)'
_FADED_LINE = 'rgba(190,190,190,0.5)'

_HIGHLIGHT_COLORS = {
    'wall_internal': ('rgba(30,110,230,0.85)', 'rgba(15,70,160,1.0)'),
    'wall_internal_estimated': ('rgba(120,180,240,0.75)', 'rgba(60,120,190,1.0)'),
    'wall_external_confirmed': ('rgba(230,90,30,0.85)', 'rgba(170,60,10,1.0)'),
    'wall_external_unknown': ('rgba(230,190,190,0.85)', 'rgba(160,50,50,1.0)'),
    'related': ('rgba(160,50,190,0.75)', 'rgba(110,20,140,1.0)'),
}
_EQUIPMENT_COLOR = 'rgba(220,190,20,0.95)'

_SPACE_FILL = 'rgba(100,180,120,0.35)'
_SPACE_FILL_SELECTED = 'rgba(230,100,60,0.55)'
_SPACE_LINE = 'rgba(60,140,80,0.9)'
_SPACE_LINE_SELECTED = 'rgba(200,60,20,1.0)'

_SPACE_UNMATCHED_FILL = 'rgba(190,190,190,0.35)'
_SPACE_UNMATCHED_LINE = 'rgba(150,150,150,0.8)'

_PAIR_PALETTE = [
    'rgba(31,119,180,0.45)', 'rgba(255,127,14,0.45)', 'rgba(44,160,44,0.45)',
    'rgba(214,39,40,0.45)', 'rgba(148,103,189,0.45)', 'rgba(140,86,75,0.45)',
    'rgba(227,119,194,0.45)', 'rgba(127,127,127,0.45)', 'rgba(188,189,34,0.45)',
    'rgba(23,190,207,0.45)',
]
_PAIR_PALETTE_LINE = [
    'rgba(31,119,180,1.0)', 'rgba(255,127,14,1.0)', 'rgba(44,160,44,1.0)',
    'rgba(214,39,40,1.0)', 'rgba(148,103,189,1.0)', 'rgba(140,86,75,1.0)',
    'rgba(227,119,194,1.0)', 'rgba(127,127,127,1.0)', 'rgba(188,189,34,1.0)',
    'rgba(23,190,207,1.0)',
]


def build_pair_labels(a_to_b):
    a_labels, b_labels = {}, {}
    for i, (a_guid, b_guid) in enumerate(a_to_b.items(), start=1):
        a_labels[a_guid] = i
        if b_guid:
            b_labels[b_guid] = i
    return a_labels, b_labels


def build_plan_figure(plan_data, click_grid_spacing=0.5, selected_guid=None,
                       highlight_map=None, equipment_entities=None, pair_labels=None):
    import plotly.graph_objects as go
    fig = go.Figure()
    highlight_map = highlight_map or {}
    pair_labels = pair_labels or {}

    # 선택된 공간의 버퍼 폴리곤 구하기 (구조 부재 클리핑 시각화 용도)
    selected_space_poly = None
    if selected_guid:
        for sp in plan_data['spaces']:
            if sp['guid'] == selected_guid:
                try:
                    selected_space_poly = sp['polygon'].buffer(_APPORTION_BUFFER_M)
                except Exception:
                    pass
                break

    # 구조요소(벽/기둥/보/바닥 등) 그리기
    for el in plan_data['structural']:
        cat = highlight_map.get(el['guid']) if highlight_map else None

        # 선택된 공간과 연관된 부재는 교차 영역만 잘라서 진하게 덧그림
        if selected_space_poly and cat in _HIGHLIGHT_COLORS:
            # 1. 전체 부재를 흐린 배경으로 먼저 그림
            xs, ys = _polygon_xy_lists(el['polygon'])
            fig.add_trace(go.Scatter(
                x=xs, y=ys, mode='lines', fill='toself',
                line=dict(width=0.5, color=_FADED_LINE), fillcolor=_FADED_COLOR,
                hoverinfo='skip', showlegend=False,
            ))
            
            # 2. 선택된 공간(버퍼 포함)과 겹치는 부분만 잘라내서 강조 색상으로 덧그림
            try:
                inter_poly = el['polygon'].intersection(selected_space_poly)
                if not inter_poly.is_empty:
                    ixs, iys = _polygon_xy_lists(inter_poly)
                    fill_c, line_c = _HIGHLIGHT_COLORS[cat]
                    fig.add_trace(go.Scatter(
                        x=ixs, y=iys, mode='lines', fill='toself',
                        line=dict(width=1.5, color=line_c), fillcolor=fill_c,
                        hoverinfo='text',
                        text=el.get('hover') or f"{el['class']} {el['name']}".strip(),
                        showlegend=False,
                    ))
            except Exception:
                pass # 클리핑 실패 시 그냥 배경만 남김
        else:
            # 선택되지 않았거나 연관 없는 부재의 일반적인 그리기
            xs, ys = _polygon_xy_lists(el['polygon'])
            if highlight_map:
                fill_c, line_c, line_w = _FADED_COLOR, _FADED_LINE, 0.5
            else:
                fill_c = _STRUCT_COLORS.get(el['class'], 'rgba(150,150,150,0.5)')
                line_c, line_w = 'rgba(60,60,60,0.6)', 0.5

            fig.add_trace(go.Scatter(
                x=xs, y=ys, mode='lines', fill='toself',
                line=dict(width=line_w, color=line_c), fillcolor=fill_c,
                hoverinfo='text',
                text=el.get('hover') or f"{el['class']} {el['name']}".strip(),
                showlegend=False,
            ))

    # Space
    badge_x, badge_y, badge_text, badge_color, badge_line = [], [], [], [], []
    for sp in plan_data['spaces']:
        is_sel = (selected_guid is not None and sp['guid'] == selected_guid)
        pair_no = pair_labels.get(sp['guid'])

        if is_sel:
            fill_c, line_c, line_w = _SPACE_FILL_SELECTED, _SPACE_LINE_SELECTED, 1.5
        elif pair_no is not None:
            idx = (pair_no - 1) % len(_PAIR_PALETTE)
            fill_c, line_c, line_w = _PAIR_PALETTE[idx], _PAIR_PALETTE_LINE[idx], 1.2
        elif pair_labels: 
            fill_c, line_c, line_w = _SPACE_UNMATCHED_FILL, _SPACE_UNMATCHED_LINE, 1.0
        else:
            fill_c, line_c, line_w = _SPACE_FILL, _SPACE_LINE, 1.0

        xs, ys = _polygon_xy_lists(sp['polygon'])
        fig.add_trace(go.Scatter(
            x=xs, y=ys, mode='lines', fill='toself',
            line=dict(width=line_w, color=line_c), fillcolor=fill_c,
            hoverinfo='skip', showlegend=False,
        ))

        if pair_no is not None:
            c = sp['polygon'].centroid
            idx = (pair_no - 1) % len(_PAIR_PALETTE)
            badge_x.append(c.x); badge_y.append(c.y)
            badge_text.append(str(pair_no))
            badge_color.append(_PAIR_PALETTE_LINE[idx])
            badge_line.append('white')

        pts = _grid_points_in_polygon(sp['polygon'], spacing=click_grid_spacing)
        gx, gy = zip(*pts)
        fig.add_trace(go.Scatter(
            x=list(gx), y=list(gy), mode='markers',
            marker=dict(size=14, opacity=0.0),
            customdata=[sp['guid']] * len(pts),
            hovertemplate=f"{sp['name']}<br>면적 약 {round(sp['polygon'].area,1)}㎡<extra></extra>",
            showlegend=False,
        ))

    if badge_x:
        fig.add_trace(go.Scatter(
            x=badge_x, y=badge_y, mode='markers+text',
            marker=dict(size=22, color=badge_color, line=dict(width=1.5, color=badge_line)),
            text=badge_text, textfont=dict(color='white', size=12, family='Arial Black'),
            hoverinfo='skip', showlegend=False,
        ))

    if equipment_entities:
        ex, ey, etext = [], [], []
        for e in equipment_entities:
            poly = get_footprint_polygon(e)
            if poly is not None and not poly.is_empty:
                c = poly.centroid
                ex.append(c.x); ey.append(c.y)
            else:
                try:
                    import ifcopenshell.util.placement as plc
                    m = plc.get_local_placement(e.ObjectPlacement)
                    ex.append(float(m[0, 3])); ey.append(float(m[1, 3]))
                except Exception:
                    continue
            etext.append(f"{e.is_a()} {e.Name or ''}".strip())
        if ex:
            fig.add_trace(go.Scatter(
                x=ex, y=ey, mode='markers',
                marker=dict(size=11, color=_EQUIPMENT_COLOR, symbol='diamond',
                            line=dict(width=1, color='rgba(120,100,0,1)')),
                text=etext, hoverinfo='text', showlegend=False,
            ))

    fig.update_xaxes(showgrid=False, zeroline=False, visible=False)
    fig.update_yaxes(showgrid=False, zeroline=False, visible=False, scaleanchor='x', scaleratio=1)
    fig.update_layout(
        margin=dict(l=10, r=10, t=30, b=10),
        height=600,
        plot_bgcolor='white',
        clickmode='event+select',
    )
    return fig