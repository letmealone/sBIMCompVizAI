# -*- coding: utf-8 -*-
"""
floorplan_app.py
------------------
전문가 IFC와 AI 생성 IFC를 업로드해 층별 평면도를 좌/우로 비교하는 Streamlit 앱.

핵심 기능:
- 층별 평면도 좌/우 비교, 공간(Space) 클릭시 접한 구조재/내외벽/면적/설비 정보 표시
- 공간 클릭시 내부(파랑)/외부(주황)/기타관련부재(보라)/설비(노랑) 색상 하이라이트
- (선택) 공간 자동 매핑: 면적+centroid 좌표 오차 임계값 기준으로 한쪽 클릭시 반대편도 자동 선택
- 비교 테이블(구조재개수/내외부구분/유형별면적/설비개수)은 두 IFC의 합집합 키로 통일해 표시

실행:
    streamlit run floorplan_app.py

요구사항: streamlit>=1.35 (st.plotly_chart의 on_select 기능 필요), plotly, shapely,
ifcopenshell, numpy, pandas, openpyxl (ifc_to_excel.py 의존성)
"""
import os
import re
import tempfile
import hashlib
import zipfile

import streamlit as st

import floorplan_core as fc
import llm_storey_match as lsm
import ifc_to_excel as ite
import comparison_export as cmpexp

st.set_page_config(layout='wide', page_title='IFC 평면도 비교')
st.title('IFC 평면도 비교 (전문가 vs AI)')

# 평면도(Plotly)에서 클릭으로 선택 가능한 지점(공간 내부 격자 마커) 위에 마우스를 올리면
# 커서가 손가락 모양(pointer)으로 바뀌도록 CSS를 주입한다 - 기본적으로 Plotly의 markers
# trace는 selectable 상태에서 보통 자체적으로 pointer 커서를 적용하지만, 브라우저/버전에
# 따라 적용이 안 되는 경우가 있어 명시적으로 강제한다. 격자 마커는 opacity=0(안 보임)이라
# 실제로 눈에 보이는 반응은 없지만, 마우스가 그 지점 위에 있을 때 커서만 바뀐다.
st.markdown(
    """
    <style>
    .js-plotly-plot .scatterlayer .trace .points path,
    .js-plotly-plot .scatterlayer .trace .points circle {
        cursor: pointer !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

MIN_STREAMLIT_VERSION = (1, 35)

def _check_streamlit_version():
    try:
        parts = tuple(int(x) for x in st.__version__.split('.')[:2])
        if parts < MIN_STREAMLIT_VERSION:
            st.warning(
                f"현재 Streamlit 버전 {st.__version__}. 평면도 클릭 선택 기능은 "
                f"Streamlit {'.'.join(map(str, MIN_STREAMLIT_VERSION))} 이상이 필요합니다. "
                f"`pip install --upgrade streamlit`을 권장합니다."
            )
    except Exception:
        pass

_check_streamlit_version()


def _extract_customdata_guid(point):
    """selection point dict에서 customdata(공간 GlobalId) 안전하게 추출."""
    cd = point.get('customdata')
    if cd is None:
        return None
    if isinstance(cd, (list, tuple)):
        return cd[0] if cd else None
    return cd


_SESSION_CACHE_PREFIX = '_cache__'

def _safe_filename_stem(filename):
    """업로드된 파일명에서 확장자를 떼고, 파일시스템/zip에 안전하지 않은 문자를 치환한다.
    엑셀/zip 출력 파일명에 원본 IFC 파일명을 그대로 반영하기 위한 용도."""
    stem = os.path.splitext(filename or '')[0]
    stem = re.sub(r'[\\/:*?"<>|]', '_', stem).strip()
    return stem or 'unnamed'


def _session_cache(key, compute_fn):
    full_key = _SESSION_CACHE_PREFIX + key
    if full_key not in st.session_state:
        st.session_state[full_key] = compute_fn()
    return st.session_state[full_key]


def _load_ifc_cached(file_bytes, filename, file_hash):
    def _compute():
        suffix = os.path.splitext(filename)[1] or '.ifc'
        with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
            tmp.write(file_bytes)
            path = tmp.name
        return fc.load_ifc(path)
    with st.spinner('IFC 파일 파싱 중... (지붕/천장 면적 등 지오메트리 계산 포함, 수 초~수십 초 소요될 수 있음)'):
        return _session_cache(f'ifc_{file_hash}', _compute)


def _precompute_storey_geometry_cached(data, file_hash, label):
    """층별 구조부재 footprint(공간-부재 지오메트리 보강탐지에 쓰임)를 파일 업로드 직후
    한 번만 미리 계산해둔다. _session_cache로 '이번 세션에 이미 실행했는지'만 체크하고,
    실제 계산 결과는 fc.precompute_storey_geometry()가 채우는 모듈 전역 캐시
    (floorplan_core._storey_candidate_footprint_cache)에 저장된다 - 이후 공간을 클릭할
    때마다 그 층을 처음 조회하며 발생하던 지연이 없어지고 캐시만 즉시 불러오게 된다."""
    def _compute():
        progress = st.progress(0.0, text=f'{label}: 층별 부재 지오메트리 사전계산 중...')
        total = len(data['storeys'])

        def _cb(storey_name, i, total_n):
            progress.progress(i / total_n, text=f'{label}: {storey_name} 사전계산 중 ({i}/{total_n})')

        fc.precompute_storey_geometry(data['ifc_file'], data['storeys'], status_cb=_cb)
        progress.empty()
        return True
    return _session_cache(f'storey_geom_precomputed_{file_hash}', _compute)


def _build_plan_cached(storeys, storey_name, cache_tag, wall_classification=None):
    def _compute():
        storey = next(s for s in storeys if s['Name'] == storey_name)
        return fc.build_storey_plan_data(storey, wall_classification=wall_classification)
    with st.spinner('해당 층 도면 지오메트리 계산 중...'):
        return _session_cache(f'plan_{cache_tag}_{storey_name}', _compute)


def _match_spaces_cached(spaces_a, spaces_b, area_thresh, centroid_thresh, cache_tag):
    def _compute():
        return fc.match_spaces(spaces_a, spaces_b, area_thresh=area_thresh, centroid_thresh=centroid_thresh)
    with st.spinner('공간 자동 매핑 계산 중... (면적으로 좌표계 오프셋 추정 후 centroid 매칭)'):
        return _session_cache(f'spacematch_{cache_tag}_{area_thresh}_{centroid_thresh}', _compute)


def _match_storeys_llm_cached(storeys_a, storeys_b, offset_mm, elevation_hint_mapping,
                               api_key, model_name, cache_tag, force_recompute=False,
                               stream_placeholder=None, space_summary_a=None, space_summary_b=None):
    key = f'llm_floor_{cache_tag}_{model_name}'
    full_key = _SESSION_CACHE_PREFIX + key
    if force_recompute and full_key in st.session_state:
        del st.session_state[full_key]

    def _on_chunk(text_so_far):
        if stream_placeholder is not None:
            stream_placeholder.code(text_so_far, language='json')

    def _compute():
        return lsm.match_storeys_llm(
            storeys_a, storeys_b, offset_mm, api_key,
            model_name=model_name, elevation_hint_mapping=elevation_hint_mapping,
            on_chunk=_on_chunk if stream_placeholder is not None else None,
            space_summary_a=space_summary_a, space_summary_b=space_summary_b,
        )
    with st.spinner(f'Gemini({model_name})로 층 매핑 중... (API 호출 1회)'):
        return _session_cache(key, _compute)


def _clear_all_caches():
    for key in list(st.session_state.keys()):
        if key.startswith((_SESSION_CACHE_PREFIX, 'left_', 'right_', '_last_storey_pair', '_file_hash_')):
            del st.session_state[key]


_FLOOR_BADGE_EMOJIS = ['🟦', '🟩', '🟧', '🟪', '🟥', '🟫', '🟨', '⬛']

def _invert_mapping(mapping):
    return {b: a for a, b in mapping.items() if b}

def _floor_pair_badges(mapping):
    a_badges, b_badges = {}, {}
    idx = 0
    for a_name, b_name in mapping.items():
        if not b_name:
            continue
        emoji = _FLOOR_BADGE_EMOJIS[idx % len(_FLOOR_BADGE_EMOJIS)]
        a_badges[a_name] = emoji
        b_badges[b_name] = emoji
        idx += 1
    return a_badges, b_badges


def _render_floor_checkbox_tree(storeys, session_selected_key, key_prefix, badges=None,
                                 mapping_to_other=None, other_session_key=None,
                                 other_sync_flag_key=None):
    valid_names = [s['Name'] for s in storeys]
    current = st.session_state.get(session_selected_key)
    if current not in valid_names:
        current = valid_names[0] if valid_names else None
        st.session_state[session_selected_key] = current

    sync_flag_key = f'{key_prefix}__sync_pending'
    if st.session_state.pop(sync_flag_key, False):
        for s in storeys:
            st.session_state[f'{key_prefix}_{s["Name"]}'] = (s['Name'] == current)

    checked_states = {}
    for s in storeys:
        name = s['Name']
        cb_key = f'{key_prefix}_{name}'
        if cb_key not in st.session_state:
            st.session_state[cb_key] = (name == current)

        badge = (badges or {}).get(name, '⬜')
        elev_text = f"{s['Elevation']:.0f}mm" if s['Elevation'] is not None else '-'
        checked_states[name] = st.checkbox(f"{badge} {name}  _(고도 {elev_text})_", key=cb_key)

    newly_checked = [n for n, c in checked_states.items() if c and n != current]
    if newly_checked:
        new_name = newly_checked[0]
        st.session_state[session_selected_key] = new_name
        st.session_state[sync_flag_key] = True
        if mapping_to_other is not None and other_session_key is not None:
            other_name = mapping_to_other.get(new_name)
            if other_name:
                st.session_state[other_session_key] = other_name
                if other_sync_flag_key:
                    st.session_state[other_sync_flag_key] = True
        st.rerun()
    elif checked_states.get(current) is False:
        st.session_state[sync_flag_key] = True
        st.rerun()

    return current


def _render_plot_and_get_detail(label, data, storey_name, plan, session_prefix, pair_labels=None,
                                 active_categories=None):
    st.subheader(label)
    if storey_name is None or plan is None:
        st.info('이 층에 대응하는 층을 찾지 못했습니다 (층 매핑 없음).')
        return None, None

    st.caption(f"공간 {len(plan['spaces'])}개 · 구조요소 {len(plan['structural'])}개  (층: {storey_name})")
    st.caption(
        '🖱️ **초록색(또는 매칭 배지 색)으로 채워진 공간 영역 안쪽 아무 곳이나 클릭**하면 선택됩니다. '
        '벽/기둥 등 회색 부재는 클릭이 아니라 **마우스를 올리면(hover)** 치수·면적·재질 정보가 보입니다. '
        '클릭이 잘 안 되면 아래 드롭다운에서 직접 선택하세요.'
    )

    selected_key = f'{session_prefix}_selected_guid'
    selected_guid = st.session_state.get(selected_key)

    dropdown_key = f'{session_prefix}_space_dropdown'
    dropdown_sync_key = f'{session_prefix}_dropdown_sync_pending'
    space_options = [''] + [s['guid'] for s in plan['spaces']]
    guid_to_name = {s['guid']: s['name'] for s in plan['spaces']}
    guid_to_area = {s['guid']: s['polygon'].area for s in plan['spaces']}

    def _fmt_space_option(guid):
        if not guid:
            return '(선택 안 함)'
        prefix = f"[{pair_labels[guid]}번] " if pair_labels and guid in pair_labels else ''
        return f"{prefix}{guid_to_name.get(guid, '?')} ({guid_to_area.get(guid, 0):.1f}㎡)"

    # 주의(버그 수정 이력): 예전엔 이 동기화를 매 실행마다 무조건 했었는데, 그러면 사용자가
    # 드롭다운에서 방금 고른 값을 위젯이 화면에 렌더링하기도 전에 selected_guid(아직 갱신
    # 되기 전의 이전 값)로 덮어써버려 "드롭다운이 반응 안 하는" 것처럼 보이는 버그가 있었다.
    # 그래서 지금은 '우리가 selected_guid를 직접 바꾸고 st.rerun()을 요청한 바로 다음 실행'
    # 에서만(dropdown_sync_key로 표시) 강제 동기화하고, 그 외의 일반 실행(=드롭다운 자체를
    # 막 조작한 경우 포함)에서는 위젯이 사용자의 선택을 그대로 유지하게 둔다.
    if st.session_state.pop(dropdown_sync_key, False):
        st.session_state[dropdown_key] = selected_guid or ''

    dropdown_guid = st.selectbox(
        '공간 직접 선택', space_options, format_func=_fmt_space_option, key=dropdown_key,
        help='평면도 클릭이 잘 안 될 때 여기서 바로 선택할 수 있습니다. 자동매핑이 켜져있으면 '
             '번호가 매겨진 항목이 반대편과 매칭된 공간입니다.',
    )

    detail = None
    sp_entry = None
    if selected_guid:
        sp_entry = next((s for s in plan['spaces'] if s['guid'] == selected_guid), None)
        if sp_entry is not None:
            detail = fc.build_space_detail(data['ifc_file'], data['wall_classification'], sp_entry['entity'])

    equipment_entities = None
    highlight_map = None
    wall_segments = None
    if detail is not None and sp_entry is not None:
        highlight_map = detail['highlight_map']
        wall_segments = detail.get('wall_segment_polygons')
        equipment_entities = fc.get_space_contained_equipment(data['ifc_file'], sp_entry['entity'])

        stats = detail.get('wall_segment_stats') or {}
        precise_cg = stats.get('precise_cg', 0)
        precise_edge = stats.get('precise_edge', 0)
        fallback_n = stats.get('fallback', 0)
        if precise_cg + precise_edge + fallback_n > 0:
            st.caption(
                f"🧱 이 공간에 접한 벽 표시 방식: **정밀(경계좌표) {precise_cg}개** / "
                f"**정밀(edge겹침 추정) {precise_edge}개** / "
                f"**전체표시(폴백) {fallback_n}개**"
            )

    try:
        fig = fc.build_plan_figure(
            plan, selected_guid=selected_guid,
            highlight_map=highlight_map, equipment_entities=equipment_entities,
            pair_labels=pair_labels, wall_segments=wall_segments,
            active_categories=active_categories,
        )
    except TypeError:
        # floorplan_core.py가 구버전이라 active_categories 파라미터 자체가 없는 경우
        # (배포시 floorplan_app.py만 최신이고 floorplan_core.py는 예전 버전인 상황) -
        # 크래시 대신 그 파라미터 없이 재시도해 범례 필터만 비활성화된 채로 동작하게 한다.
        fig = fc.build_plan_figure(
            plan, selected_guid=selected_guid,
            highlight_map=highlight_map, equipment_entities=equipment_entities,
            pair_labels=pair_labels, wall_segments=wall_segments,
        )
    event = st.plotly_chart(
        fig, key=f'{session_prefix}_plot', on_select='rerun',
        selection_mode=('points',), use_container_width=True,
    )

    new_guid = None
    if dropdown_guid and dropdown_guid != selected_guid:
        new_guid = dropdown_guid
    elif event and event.get('selection', {}).get('points'):
        g = _extract_customdata_guid(event['selection']['points'][0])
        if g and g != selected_guid:
            new_guid = g

    if detail is not None:
        badge = f" `[{pair_labels[detail['guid']]}번]`" if pair_labels and detail['guid'] in pair_labels else ''
        st.markdown(f"**📍 {detail['name']}**{badge}" + (f" ({detail['long_name']})" if detail['long_name'] else ''))
        c1, c2 = st.columns(2)
        with c1:
            st.metric('공간 면적(㎡)', detail['area'] if detail['area'] is not None else 'N/A')
            st.caption(f"산출방식: {detail['area_method']}")
        with c2:
            st.caption(f"GlobalId: `{detail['guid']}`")
    elif selected_guid:
        st.warning('선택된 공간을 이 층에서 찾을 수 없습니다 (층이 바뀌었을 수 있음).')
    else:
        st.caption('평면도에서 공간을 클릭하거나, 위 드롭다운에서 선택하면 상세 정보가 여기 표시됩니다.')

    return detail, new_guid


def _render_legend_filter():
    """평면도 범례를 인터랙티브 필터로 렌더링한다. 벽 이외 부재/설비를 하나로 뭉뚱그리지
    않고 개별 클래스별 항목으로 나열하며, 선택 해제한 항목은 양쪽 평면도 모두에서
    옅게(배경처럼) 처리된다(기본은 전체 선택 = 기존과 동일하게 전부 강조 표시).
    반환: 현재 활성화된 카테고리 집합(set), 또는 floorplan_core.py가 구버전이라
    get_legend_items가 없으면 None(필터 없이 기존 방식대로 전체 강조 표시)."""
    get_items_fn = getattr(fc, 'get_legend_items', None)
    if get_items_fn is None:
        st.caption(
            '🎨 범례: 내부/외부 벽 + 관련부재 + 설비 (구버전 floorplan_core.py가 배포되어 '
            '개별 항목 필터는 비활성화됨 - floorplan_core.py를 최신본으로 교체해주세요)'
        )
        return None

    items = get_items_fn()
    options = [key for key, _label, _color in items]
    labels = {key: label for key, label, _color in items}

    key = '_legend_active_categories'
    if key not in st.session_state:
        st.session_state[key] = options  # 최초 진입시 기본값: 전체 선택

    selected = st.multiselect(
        '🎨 범례 (선택된 항목만 평면도에서 강조 표시됩니다 · 기본은 전체 선택)',
        options=options, format_func=lambda k: labels.get(k, k), key=key,
    )
    return set(selected)


def _render_union_table(title, left_d, right_d, label_left, label_right,
                         extra_left=None, extra_right=None, extra_label=None):
    if not left_d and not right_d:
        return
    keys = list(dict.fromkeys(list(left_d.keys()) + list(right_d.keys())))
    st.markdown(f'**{title}**')
    table = {
        '구분': keys,
        label_left: [left_d.get(k, 0) for k in keys],
        label_right: [right_d.get(k, 0) for k in keys],
    }
    if extra_left is not None and extra_right is not None:
        table[f'{label_left}·{extra_label}'] = [extra_left.get(k, 0) for k in keys]
        table[f'{label_right}·{extra_label}'] = [extra_right.get(k, 0) for k in keys]
    st.table(table)


def _render_comparison_tables(detail_left, detail_right, label_left='전문가', label_right='AI'):
    if detail_left is None and detail_right is None:
        return
    st.markdown('---')
    st.markdown('## 🔍 선택된 공간 비교')

    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f"**{label_left}**: " + (f"{detail_left['name']} (면적 {detail_left['area']}㎡)"
                    if detail_left else '선택된 공간 없음'))
    with c2:
        st.markdown(f"**{label_right}**: " + (f"{detail_right['name']} (면적 {detail_right['area']}㎡)"
                    if detail_right else '선택된 공간 없음'))

    dl = detail_left or {}
    dr = detail_right or {}

    _render_union_table('접한 구조재 개수', dl.get('class_counts', {}), dr.get('class_counts', {}),
                         label_left, label_right)

    _render_union_table('벽 내부/외부 구분', dl.get('wall_simple_counts', {}), dr.get('wall_simple_counts', {}),
                         label_left, label_right,
                         extra_left=dl.get('wall_simple_area', {}), extra_right=dr.get('wall_simple_area', {}),
                         extra_label='합산면적(㎡)')
    st.caption(
        f"💡 벽 면적은 여러 공간에 걸친 벽의 경우 이 공간에 해당하는 부분만 안분해서 합산합니다 "
        f"({label_left}: {dl.get('wall_area_note', '')} · {label_right}: {dr.get('wall_area_note', '')})"
    )

    area_l = {k: v['면적합계(㎡)'] for k, v in dl.get('area_by_class', {}).items()}
    area_r = {k: v['면적합계(㎡)'] for k, v in dr.get('area_by_class', {}).items()}
    _render_union_table('벽 이외 부재 유형별 합산 면적(㎡)', area_l, area_r, label_left, label_right)
    apportion_notes = {k: v['비고'] for k, v in dl.get('area_by_class', {}).items() if v.get('비고')}
    apportion_notes.update({k: v['비고'] for k, v in dr.get('area_by_class', {}).items() if v.get('비고')})
    if apportion_notes:
        st.caption('💡 바닥/지붕/천장(IfcSlab/IfcRoof/IfcCovering)도 여러 공간에 걸친 경우 이 공간 몫만 안분: '
                   + ', '.join(f'{k}{v}' for k, v in apportion_notes.items()))

    _render_union_table('설비 개수', dl.get('equipment_counts', {}), dr.get('equipment_counts', {}),
                         label_left, label_right)


# ===================================================================
# 파일 업로드 (메인 화면)
# ===================================================================

col_up1, col_up2 = st.columns(2)
with col_up1:
    file_a = st.file_uploader('전문가 IFC 업로드', type=['ifc'], key='upload_a')
with col_up2:
    file_b = st.file_uploader('AI 생성 IFC 업로드', type=['ifc'], key='upload_b')

if file_a and file_b:
    file_hash_a = hashlib.md5(file_a.getvalue()).hexdigest()[:10]
    file_hash_b = hashlib.md5(file_b.getvalue()).hexdigest()[:10]

    prev_hash_a = st.session_state.get('_file_hash_a')
    prev_hash_b = st.session_state.get('_file_hash_b')
    if (prev_hash_a is not None and prev_hash_a != file_hash_a) or \
       (prev_hash_b is not None and prev_hash_b != file_hash_b):
        _clear_all_caches()
        st.toast('새 IFC 파일이 감지되어 이전 캐시를 자동으로 비웠습니다.', icon='🔄')
    st.session_state['_file_hash_a'] = file_hash_a
    st.session_state['_file_hash_b'] = file_hash_b

    data_a = _load_ifc_cached(file_a.getvalue(), file_a.name, file_hash_a)
    data_b = _load_ifc_cached(file_b.getvalue(), file_b.name, file_hash_b)

    _precompute_storey_geometry_cached(data_a, file_hash_a, '전문가 IFC')
    _precompute_storey_geometry_cached(data_b, file_hash_b, 'AI IFC')

    floor_mapping, floor_offset = fc.match_storeys(data_a['storeys'], data_b['storeys'])

    # ===================================================================
    # 사이드바 렌더링 (순서: 1.초기화 -> 2.AI 층 매핑 -> 3.공간 자동 매핑 -> 4.층 선택 -> 5.엑셀 내보내기)
    # ===================================================================
    with st.sidebar:
        # ---------------------------------------------------------------
        # [1] 초기화
        # ---------------------------------------------------------------
        st.header('🔄 초기화')
        if st.button('내 세션 캐시 초기화 (문제 있을 때)', width='stretch'):
            _clear_all_caches()
            st.success('이 세션의 캐시를 초기화했습니다. (다른 사용자에게는 영향 없음)')
            st.rerun()
        st.caption(
            '평면도가 이전에 올린 파일 내용처럼 보이는 등 문제가 있을 때 눌러주세요. '
            '새 IFC를 업로드하면 자동으로도 초기화됩니다.'
        )
        st.divider()

        # ---------------------------------------------------------------
        # [2] AI 층 매핑 (Gemini)
        # ---------------------------------------------------------------
        st.header('🤖 AI 층 매핑 (Gemini)')
        ai_floor_map_enabled = st.checkbox(
            '층 이름의 문맥적 연관성 + 표고 유사도 + 공간(Space) 구성 유사도를 함께 고려해 AI로 층 매핑',
            value=False
        )
        
        llm_floor_mapping, llm_floor_detail = None, None
        
        if ai_floor_map_enabled:
            _secret_key = st.secrets.get('GOOGLE_API_KEY', '')
            if _secret_key:
                google_api_key = _secret_key
                st.caption('✅ `.streamlit/secrets.toml`의 `GOOGLE_API_KEY`를 사용합니다.')
            else:
                google_api_key = st.text_input(
                    'Google AI API Key', type='password',
                    value=st.session_state.get('_google_api_key', '')
                )
                st.session_state['_google_api_key'] = google_api_key
                
            llm_model_name = st.selectbox(
                '모델', options=['gemini-3.1-flash-lite', 'gemini-2.5-flash-lite', 'gemini-2.5-flash'],
                index=0
            )
            run_llm_floor_map = st.button('🚀 AI 층 매핑 실행 (API 1회 호출)', width='stretch')
            st.caption('⚠️ 이 버튼을 누를 때만 API가 호출됩니다. 체크박스만 켜는 것으로는 호출되지 않습니다.')

            if not google_api_key:
                st.warning('AI 층 매핑을 실행하려면 Google AI API Key를 입력하세요.')
            else:
                _llm_cache_tag = f'{file_hash_a}_{file_hash_b}'
                _llm_full_key = _SESSION_CACHE_PREFIX + f'llm_floor_{_llm_cache_tag}_{llm_model_name}'
                
                _stream_area = None
                if run_llm_floor_map:
                    st.markdown('**🔎 Gemini 실시간 추론 진행**')
                    _stream_area = st.empty()
                    _stream_area.code('(응답 대기 중...)', language='json')

                if run_llm_floor_map or _llm_full_key in st.session_state:
                    try:
                        _space_summary_a = {s['Name']: fc.get_storey_space_summary(s) for s in data_a['storeys']}
                        _space_summary_b = {s['Name']: fc.get_storey_space_summary(s) for s in data_b['storeys']}
                        llm_floor_mapping, llm_floor_detail = _match_storeys_llm_cached(
                            data_a['storeys'], data_b['storeys'], floor_offset, floor_mapping,
                            google_api_key, llm_model_name, _llm_cache_tag,
                            force_recompute=run_llm_floor_map, stream_placeholder=_stream_area,
                            space_summary_a=_space_summary_a, space_summary_b=_space_summary_b,
                        )
                        if _stream_area is not None:
                            st.caption('✅ 파싱 결과가 아래 층 선택 배지에 반영되었습니다.')
                    except lsm.LlmStoreyMatchError as e:
                        st.error(f'AI 층 매핑 실패: {e}')

        # AI 매핑 결과 상세 정보 표시 위치 이동
        if llm_floor_mapping is not None:
            st.caption('같은 색 배지 = AI가 매핑한 층입니다. ⬜는 대응되는 층을 못 찾은 경우입니다.')
            with st.expander('AI 매핑 상세 근거 보기'):
                for d in llm_floor_detail:
                    st.caption(f"`{d['a_name']}` → `{d['b_name']}` (확신도: {d['confidence']}) — {d['reason']}")
        
        st.divider()

        # ---------------------------------------------------------------
        # [3] 공간 자동 매핑
        # ---------------------------------------------------------------
        st.header('⚙️ 공간 자동 매핑')
        # 상태 유지를 위해 key 지정, 기본값은 True (Checked)
        auto_map_enabled = st.checkbox(
            '면적 + centroid 좌표 오차 기준으로 자동 매핑', 
            value=True, 
            key='auto_space_map_cb'
        )
        if auto_map_enabled:
            area_thresh = st.number_input('면적 오차 임계값 (㎡)', min_value=0.0, value=2.0, step=0.5)
            centroid_thresh = st.number_input('centroid 좌표 오차 임계값 (m)', min_value=0.0, value=1.0, step=0.1)
        else:
            area_thresh = centroid_thresh = None

        st.divider()

        # ---------------------------------------------------------------
        # [4] 층 선택
        # ---------------------------------------------------------------
        st.header('🏢 층 선택')
        
        _badge_source = llm_floor_mapping if llm_floor_mapping is not None else floor_mapping
        _badge_source_inv = _invert_mapping(_badge_source)
        floor_badges_a, floor_badges_b = _floor_pair_badges(_badge_source)

        floor_auto_follow = st.checkbox('매핑된 층 자동 함께 선택', value=True)
        if llm_floor_mapping is None:
            st.caption('같은 색 배지가 붙은 층끼리 고도 기준 자동 매핑된 층입니다. ⬜는 대응되는 층을 못 찾은 경우입니다.')

        with st.expander('전문가 IFC', expanded=True):
            selected_a_name = _render_floor_checkbox_tree(
                data_a['storeys'], 'left_selected_floor', 'left_floor_cb', badges=floor_badges_a,
                mapping_to_other=(_badge_source if floor_auto_follow else None),
                other_session_key='right_selected_floor',
                other_sync_flag_key='right_floor_cb__sync_pending')
        with st.expander('AI IFC', expanded=True):
            selected_b_name = _render_floor_checkbox_tree(
                data_b['storeys'], 'right_selected_floor', 'right_floor_cb', badges=floor_badges_b,
                mapping_to_other=(_badge_source_inv if floor_auto_follow else None),
                other_session_key='left_selected_floor',
                other_sync_flag_key='left_floor_cb__sync_pending')

        st.divider()

        # ---------------------------------------------------------------
        # [5] 엑셀 내보내기
        # ---------------------------------------------------------------
        st.header('📊 엑셀 내보내기')
        st.caption('전문가/AI IFC 각각의 전체 부재 추출 결과와, 층/공간 매칭·객체 비교 결과를 엑셀 3개로 만들어 zip 하나로 묶어 내려받습니다.')
        export_area_thresh = st.number_input(
            '내보내기용 면적 오차 임계값 (㎡)', min_value=0.0,
            value=area_thresh if area_thresh else 2.0, step=0.5, key='export_area_thresh',
        )
        export_centroid_thresh = st.number_input(
            '내보내기용 centroid 오차 임계값 (m)', min_value=0.0,
            value=centroid_thresh if centroid_thresh else 1.0, step=0.1, key='export_centroid_thresh',
        )
        
        # 엑셀 내보내기 버튼
        run_export = st.button('📥 비교 엑셀(zip) 생성', width='stretch')

        # --- 엑셀 생성 실행 블록 ---
        if run_export:
            export_status = st.empty()

            def _export_status_cb(msg):
                export_status.caption(f'⏳ {msg}')

            try:
                export_status.caption('⏳ 렌더링된 공간 지오메트리 캐시 확인 및 추출 중...')
                
                spaces_dict_a = {}
                for s in data_a['storeys']:
                    name = s['Name']
                    cache_key = f"{_SESSION_CACHE_PREFIX}plan_left_{file_hash_a}_{name}"
                    if cache_key in st.session_state:
                        spaces_dict_a[name] = st.session_state[cache_key]['spaces']
                    else:
                        spaces_dict_a[name] = cmpexp._floor_space_polygons(s)

                spaces_dict_b = {}
                for s in data_b['storeys']:
                    name = s['Name']
                    cache_key = f"{_SESSION_CACHE_PREFIX}plan_right_{file_hash_b}_{name}"
                    if cache_key in st.session_state:
                        spaces_dict_b[name] = st.session_state[cache_key]['spaces']
                    else:
                        spaces_dict_b[name] = cmpexp._floor_space_polygons(s)

                a_stem = _safe_filename_stem(file_a.name)
                b_stem = _safe_filename_stem(file_b.name)

                with st.spinner('전문가(A) IFC 전체 부재 엑셀 생성 중...'):
                    _tmp_a_path = tempfile.NamedTemporaryFile(suffix='.ifc', delete=False).name
                    with open(_tmp_a_path, 'wb') as fpa:
                        fpa.write(file_a.getvalue())
                    out_a = ite.extract_ifc_to_excel(
                        _tmp_a_path, os.path.join(tempfile.gettempdir(), f'{a_stem}_추출_{file_hash_a}.xlsx'))

                with st.spinner('AI(B) IFC 전체 부재 엑셀 생성 중...'):
                    _tmp_b_path = tempfile.NamedTemporaryFile(suffix='.ifc', delete=False).name
                    with open(_tmp_b_path, 'wb') as fpb:
                        fpb.write(file_b.getvalue())
                    out_b = ite.extract_ifc_to_excel(
                        _tmp_b_path, os.path.join(tempfile.gettempdir(), f'{b_stem}_추출_{file_hash_b}.xlsx'))

                wb_compare = cmpexp.build_comparison_workbook(
                    data_a, data_b, _badge_source,
                    spaces_dict_a=spaces_dict_a, spaces_dict_b=spaces_dict_b,
                    is_llm_floor_mapping=(llm_floor_mapping is not None), llm_floor_detail=llm_floor_detail,
                    area_thresh=export_area_thresh, centroid_thresh=export_centroid_thresh,
                    status_cb=_export_status_cb,
                )
                out_compare = os.path.join(
                    tempfile.gettempdir(), f'{a_stem}_{b_stem}_비교결과_{file_hash_a}_{file_hash_b}.xlsx')
                wb_compare.save(out_compare)
                export_status.empty()

                zip_filename = f'{a_stem}_{b_stem}_IFC비교.zip'
                zip_path = os.path.join(tempfile.gettempdir(), f'{a_stem}_{b_stem}_{file_hash_a}_{file_hash_b}.zip')
                with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                    zf.write(out_a, arcname=f'01_{a_stem}_추출.xlsx')
                    zf.write(out_b, arcname=f'02_{b_stem}_추출.xlsx')
                    zf.write(out_compare, arcname=f'03_{a_stem}_{b_stem}_비교결과.xlsx')
                with open(zip_path, 'rb') as fz:
                    st.session_state['_export_zip_bytes'] = fz.read()
                st.session_state['_export_zip_filename'] = zip_filename
                st.success('✅ 엑셀 생성 완료!')
            except Exception as e:
                st.error(f'엑셀 생성 중 오류가 발생했습니다: {e}')

        # --- 다운로드 버튼 (엑셀 생성 실행 블록 바로 아래) ---
        if st.session_state.get('_export_zip_bytes'):
            st.download_button(
                '⬇️ 비교 결과 zip 다운로드', data=st.session_state['_export_zip_bytes'],
                file_name=st.session_state.get('_export_zip_filename', 'IFC_비교결과.zip'),
                mime='application/zip', width='stretch',
            )
        # ===================================================================

    # 층 선택이 바뀌면 이전 선택된 공간 정보는 초기화
    _cur_key = (selected_a_name, selected_b_name)
    if st.session_state.get('_last_storey_pair') != _cur_key:
        st.session_state.pop('left_selected_guid', None)
        st.session_state.pop('right_selected_guid', None)
        st.session_state.pop('left_space_dropdown', None)
        st.session_state.pop('right_space_dropdown', None)
        st.session_state.pop('left_dropdown_sync_pending', None)
        st.session_state.pop('right_dropdown_sync_pending', None)
        st.session_state['_last_storey_pair'] = _cur_key

    st.markdown(f"### 비교 중: 전문가 `{selected_a_name}` ↔ AI `{selected_b_name}`")
    st.caption('층은 왼쪽 사이드바 "🏢 층 선택"에서 바꿀 수 있습니다.')

    plan_a = _build_plan_cached(data_a['storeys'], selected_a_name, f'left_{file_hash_a}',
                                 wall_classification=data_a['wall_classification'])
    plan_b = _build_plan_cached(data_b['storeys'], selected_b_name, f'right_{file_hash_b}',
                                 wall_classification=data_b['wall_classification'])

    space_a_to_b, space_b_to_a = {}, {}
    pair_labels_a, pair_labels_b = None, None
    if auto_map_enabled:
        space_a_to_b, space_b_to_a, match_offset, match_info = _match_spaces_cached(
            plan_a['spaces'], plan_b['spaces'], area_thresh, centroid_thresh,
            f'{file_hash_a}_{selected_a_name}|{file_hash_b}_{selected_b_name}',
        )
        if match_offset:
            pair_labels_a, pair_labels_b = fc.build_pair_labels(space_a_to_b)
            st.success(
                f"공간 자동 매핑: {len(match_info)}쌍 매칭됨 (평면도에 같은 번호·색상으로 표시됩니다) "
                f"· 추정 좌표 오프셋 dx={match_offset[0]:.2f}m, dy={match_offset[1]:.2f}m"
            )
            st.caption(
                '평면도의 색칠된 숫자 배지는 매칭된 공간 쌍입니다 (양쪽에서 같은 번호=같은 색). '
                '⬜ 회색은 반대편에서 대응되는 공간을 찾지 못한 경우입니다.'
            )
        else:
            st.warning('공간 자동 매핑: 매칭 후보를 찾지 못했습니다 (면적 임계값을 늘려보세요).')

    active_categories = _render_legend_filter()

    col_left, col_right = st.columns(2)
    with col_left:
        detail_left, new_left = _render_plot_and_get_detail(
            '전문가 IFC', data_a, selected_a_name, plan_a, 'left', pair_labels=pair_labels_a,
            active_categories=active_categories)
    with col_right:
        detail_right, new_right = _render_plot_and_get_detail(
            'AI IFC', data_b, selected_b_name, plan_b, 'right', pair_labels=pair_labels_b,
            active_categories=active_categories)

    # (범례는 위에서 _render_legend_filter()로 인터랙티브하게 이미 표시됨)

    changed = False
    if new_left:
        st.session_state['left_selected_guid'] = new_left
        st.session_state['left_dropdown_sync_pending'] = True
        changed = True
        if auto_map_enabled and new_left in space_a_to_b:
            st.session_state['right_selected_guid'] = space_a_to_b[new_left]
            st.session_state['right_dropdown_sync_pending'] = True
    if new_right:
        st.session_state['right_selected_guid'] = new_right
        st.session_state['right_dropdown_sync_pending'] = True
        changed = True
        if auto_map_enabled and new_right in space_b_to_a:
            st.session_state['left_selected_guid'] = space_b_to_a[new_right]
            st.session_state['left_dropdown_sync_pending'] = True
    if changed:
        st.rerun()  # 하이라이트/자동매핑 반영을 위해 갱신된 session_state로 즉시 재실행

    _render_comparison_tables(detail_left, detail_right, '전문가', 'AI')

else:
    st.info('좌측/우측에 전문가 IFC와 AI 생성 IFC를 각각 업로드해주세요.')
    
    with st.sidebar:
        st.header('🔄 초기화')
        st.header('🤖 AI 층 매핑 (Gemini)')
        st.header('⚙️ 공간 자동 매핑')
        st.header('🏢 층 선택')
        st.header('📊 엑셀 내보내기')
        st.caption('IFC 파일을 모두 업로드하면 세부 설정 메뉴가 활성화됩니다.')