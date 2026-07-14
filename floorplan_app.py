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
import tempfile
import hashlib

import streamlit as st

import floorplan_core as fc
import llm_storey_match as lsm

st.set_page_config(layout='wide', page_title='IFC 평면도 비교')
st.title('IFC 평면도 비교 (전문가 vs AI)')

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


def _session_cache(key, compute_fn):
    """이 세션(session_state)에만 저장되는 캐시. st.cache_resource와 달리 다른 사용자의
    세션에는 전혀 영향을 주지 않는다 - session_state 자체가 세션별로 격리되어 있기 때문.
    key가 이미 있으면 재계산 없이 그대로 반환, 없으면 compute_fn()을 실행해 저장 후 반환."""
    full_key = _SESSION_CACHE_PREFIX + key
    if full_key not in st.session_state:
        st.session_state[full_key] = compute_fn()
    return st.session_state[full_key]


def _load_ifc_cached(file_bytes, filename, file_hash):
    """업로드된 IFC를 임시파일로 저장 후 파싱. 이 세션 안에서 같은 파일(해시로 식별)에 대해
    1회만 실행되도록 session_state에 저장."""
    def _compute():
        suffix = os.path.splitext(filename)[1] or '.ifc'
        with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
            tmp.write(file_bytes)
            path = tmp.name
        return fc.load_ifc(path)
    with st.spinner('IFC 파일 파싱 중... (지붕/천장 면적 등 지오메트리 계산 포함, 수 초~수십 초 소요될 수 있음)'):
        return _session_cache(f'ifc_{file_hash}', _compute)


def _build_plan_cached(storeys, storey_name, cache_tag):
    """층별 평면 지오메트리를 이 세션 안에서만 캐싱.
    cache_tag: 파일해시 등을 포함한 문자열로, 같은 층 이름이라도 파일이 다르면 구분되게 한다."""
    def _compute():
        storey = next(s for s in storeys if s['Name'] == storey_name)
        return fc.build_storey_plan_data(storey)
    with st.spinner('해당 층 도면 지오메트리 계산 중...'):
        return _session_cache(f'plan_{cache_tag}_{storey_name}', _compute)


def _match_spaces_cached(spaces_a, spaces_b, area_thresh, centroid_thresh, cache_tag):
    """공간 자동 매핑 결과를 이 세션 안에서만 캐싱."""
    def _compute():
        return fc.match_spaces(spaces_a, spaces_b, area_thresh=area_thresh, centroid_thresh=centroid_thresh)
    with st.spinner('공간 자동 매핑 계산 중... (면적으로 좌표계 오프셋 추정 후 centroid 매칭)'):
        return _session_cache(f'spacematch_{cache_tag}_{area_thresh}_{centroid_thresh}', _compute)


def _match_storeys_llm_cached(storeys_a, storeys_b, offset_mm, elevation_hint_mapping,
                               api_key, model_name, cache_tag, force_recompute=False,
                               stream_placeholder=None):
    """AI(Gemini) 층 매핑 결과를 이 세션 안에서만 캐싱한다.
    같은 (파일쌍, 모델) 조합에 대해서는 st.rerun()이 아무리 반복돼도 API를 다시 부르지
    않는다 - force_recompute=True(실행 버튼을 눌렀을 때)일 때만 실제로 1회 호출한다.
    stream_placeholder가 주어지면(st.empty() 등) 실제로 API를 호출하는 이번 실행에 한해
    Gemini의 스트리밍 응답을 실시간으로 그 자리에 표시한다(캐시 hit일 때는 호출 자체가
    없으므로 표시할 것도 없다)."""
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
        )
    with st.spinner(f'Gemini({model_name})로 층 매핑 중... (API 호출 1회)'):
        return _session_cache(key, _compute)


def _clear_all_caches():
    """이 세션의 캐시(IFC 파싱/평면 지오메트리/공간매칭 결과)와 선택 상태를 전부 비운다.
    session_state 기반이라 다른 사용자의 세션에는 전혀 영향을 주지 않는다.
    층/공간 선택 관련 키는 전부 'left_'/'right_' 접두사로 통일해뒀기 때문에 이 한 번의
    접두사 매칭으로 다 같이 지워진다 (새 파일 업로드시 이전 파일 정보가 안 남도록 하는 핵심 장치)."""
    for key in list(st.session_state.keys()):
        if key.startswith((_SESSION_CACHE_PREFIX, 'left_', 'right_', '_last_storey_pair', '_file_hash_')):
            del st.session_state[key]


_FLOOR_BADGE_EMOJIS = ['🟦', '🟩', '🟧', '🟪', '🟥', '🟫', '🟨', '⬛']


def _invert_mapping(mapping):
    """{A이름: B이름} 매핑을 {B이름: A이름}으로 뒤집는다 (반대편 자동 선택에 사용)."""
    return {b: a for a, b in mapping.items() if b}


def _floor_pair_badges(mapping):
    """match_storeys()가 반환한 {A층이름: B층이름} 매핑으로 (a_badges, b_badges) 생성.
    같은 쌍은 항상 같은 색 이모지를 받는다 (공간 자동매핑의 번호배지와 같은 개념,
    체크박스 목록에는 배경색을 직접 칠할 수 없어 색깔 이모지로 대신한다)."""
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
    """체크박스 목록으로 층 하나를 라디오처럼(하나만) 선택하게 하는 위젯.
    session_selected_key/key_prefix는 반드시 'left_'/'right_'로 시작해야 새 파일 업로드시
    _clear_all_caches()의 접두사 매칭으로 자동 초기화된다 (기존 IFC 정보 잔존 방지).
    반환: 현재 선택된 층 이름.

    mapping_to_other/other_session_key/other_sync_flag_key: 주어지면(자동매핑 활성화시),
    이 트리에서 새 층을 선택했을 때 매핑된(mapping_to_other[선택층]) 반대편 층도 함께
    선택되도록 반대편의 session_state와 sync 플래그를 같이 갱신한다. 매핑에 대응 층이
    없으면(None) 반대편은 그대로 둔다.

    주의(버그 수정 이력): 체크박스 생성 '직전에' 매번 강제로 상태를 동기화하면, 사용자가
    방금 클릭한 값을 코드가 읽기도 전에 덮어써버려 클릭이 무시되는 문제가 있었다.
    그래서 강제 동기화는 '우리가 직접 st.rerun()을 요청한 바로 다음 실행'에서만
    (sync_flag_key로 표시) 적용하고, 그 외의 일반 실행에서는 체크박스를 있는 그대로
    두어 사용자의 클릭이 먼저 반영되도록 한다."""
    valid_names = [s['Name'] for s in storeys]
    current = st.session_state.get(session_selected_key)
    if current not in valid_names:
        current = valid_names[0] if valid_names else None
        st.session_state[session_selected_key] = current

    sync_flag_key = f'{key_prefix}__sync_pending'
    if st.session_state.pop(sync_flag_key, False):
        # 직전 실행에서 선택이 바뀌어 우리가 rerun을 요청한 그 다음 실행 -> 지금은
        # 아직 어떤 체크박스도 이번 실행에서 생성되지 않았으므로 안전하게 강제 동기화 가능
        for s in storeys:
            st.session_state[f'{key_prefix}_{s["Name"]}'] = (s['Name'] == current)

    checked_states = {}
    for s in storeys:
        name = s['Name']
        cb_key = f'{key_prefix}_{name}'
        if cb_key not in st.session_state:
            st.session_state[cb_key] = (name == current)  # 최초 생성시 기본값만 설정

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
        # 현재 선택된 항목의 체크를 사용자가 해제하려 한 경우 -> 최소 1개는 선택되어야 하므로
        # 즉시 되돌린다(라디오 버튼처럼 동작하게 하기 위함)
        st.session_state[sync_flag_key] = True
        st.rerun()

    return current


def _render_plot_and_get_detail(label, data, storey_name, plan, session_prefix, pair_labels=None):
    """평면도 렌더링 + 클릭/드롭다운 선택 처리. (선택된 공간의 detail, 새로 클릭/선택된 guid) 반환."""
    st.subheader(label)
    if storey_name is None or plan is None:
        st.info('이 층에 대응하는 층을 찾지 못했습니다 (층 매핑 없음).')
        return None, None

    st.caption(f"공간 {len(plan['spaces'])}개 · 구조요소 {len(plan['structural'])}개  (층: {storey_name})")

    selected_key = f'{session_prefix}_selected_guid'
    selected_guid = st.session_state.get(selected_key)

    # 드롭다운으로 직접 선택 (클릭이 여러 번 필요해 불편한 경우를 위한 안정적인 대안 경로.
    # 기존 클릭 방식은 그대로 두고 "추가"하는 것이라 기존 동작에는 영향이 없다)
    dropdown_key = f'{session_prefix}_space_dropdown'
    space_options = [''] + [s['guid'] for s in plan['spaces']]
    guid_to_name = {s['guid']: s['name'] for s in plan['spaces']}
    guid_to_area = {s['guid']: s['polygon'].area for s in plan['spaces']}

    def _fmt_space_option(guid):
        if not guid:
            return '(선택 안 함)'
        prefix = f"[{pair_labels[guid]}번] " if pair_labels and guid in pair_labels else ''
        return f"{prefix}{guid_to_name.get(guid, '?')} ({guid_to_area.get(guid, 0):.1f}㎡)"

    # 클릭으로 바뀐 선택과 드롭다운 위젯 상태가 항상 일치하도록 동기화
    if st.session_state.get(dropdown_key) != (selected_guid or ''):
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
    if detail is not None and sp_entry is not None:
        highlight_map = detail['highlight_map']
        equipment_entities = fc.get_space_contained_equipment(data['ifc_file'], sp_entry['entity'])

    fig = fc.build_plan_figure(
        plan, selected_guid=selected_guid,
        highlight_map=highlight_map, equipment_entities=equipment_entities,
        pair_labels=pair_labels,
    )
    event = st.plotly_chart(
        fig, key=f'{session_prefix}_plot', on_select='rerun',
        selection_mode=('points',), use_container_width=True,
    )

    # 새 선택 판단: 드롭다운에서 바뀌었으면 그걸 우선, 아니면 클릭 이벤트를 확인
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


def _render_legend():
    st.caption(
        '🟦 내부-확정(1차/2차) · 🟦(연함) 내부-추정(3차: 관계상 서로다른 Space 2개+ 연결, '
        '지오메트리 미확인) · 🟧 외부-판정됨 · 🟥 외부-판정불가(근거 없어 편입) · '
        '🟪 벽 이외 관련부재(기둥/문/창/바닥 등) · 🟨 설비(조명·센서·소방장치) · '
        '⬜ 선택된 공간과 무관한 배경 요소'
    )


def _render_union_table(title, left_d, right_d, label_left, label_right,
                         extra_left=None, extra_right=None, extra_label=None):
    """left_d/right_d(dict) 키의 합집합을 행으로 하는 통일된 비교 테이블 렌더링.
    (한쪽에만 있는 클래스도 다른 쪽엔 0으로 채워져 두 IFC 표의 행 구성이 항상 동일해진다)"""
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
    """선택된 두 공간(좌/우)의 지표를 합집합 기준 통일 테이블로 나란히 비교."""
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
# 사이드바: 초기화 + 공간 자동 매핑 설정
# ===================================================================

with st.sidebar:
    st.header('🔄 초기화')
    if st.button('내 세션 캐시 초기화 (문제 있을 때)', width='stretch'):
        _clear_all_caches()
        st.success('이 세션의 캐시를 초기화했습니다. (다른 사용자에게는 영향 없음)')
        st.rerun()
    st.caption(
        '평면도가 이전에 올린 파일 내용처럼 보이는 등 문제가 있을 때 눌러주세요. '
        '이 캐시는 세션(브라우저 탭)별로 독립되어 있어 다른 사용자의 화면에는 영향을 주지 않습니다. '
        '새 IFC를 업로드하면 자동으로도 초기화됩니다.'
    )
    st.divider()

    st.header('⚙️ 공간 자동 매핑')
    auto_map_enabled = st.checkbox(
        '면적 + centroid 좌표 오차 기준으로 자동 매핑',
        value=False,
        help='한쪽 평면도에서 공간을 클릭하면, 두 IFC의 좌표계 차이(평행이동)를 면적이 '
             '비슷한 후보들로부터 자동 추정한 뒤, 그 오프셋을 보정한 centroid 거리와 면적 오차가 '
             '둘 다 임계값 이내인 공간을 반대편에서 자동으로 찾아 함께 선택합니다.',
    )
    if auto_map_enabled:
        area_thresh = st.number_input('면적 오차 임계값 (㎡)', min_value=0.0, value=2.0, step=0.5)
        centroid_thresh = st.number_input('centroid 좌표 오차 임계값 (m)', min_value=0.0, value=1.0, step=0.1)
        st.caption(
            '⚠️ 두 모델의 좌표계가 회전 없이 평행이동만 다르다고 가정합니다. '
            '건물이 회전되어 모델링된 경우 이 방식이 맞지 않을 수 있습니다.'
        )
    else:
        area_thresh = centroid_thresh = None

    st.divider()
    st.header('🤖 AI 층 매핑 (Gemini)')
    ai_floor_map_enabled = st.checkbox(
        '층 이름의 문맥적 연관성 + 표고 유사도를 함께 고려해 AI로 층 매핑',
        value=False,
        help='표고(고도) 기반 매핑은 이미 항상 계산되어 배지로 표시되지만, 이 옵션은 '
             '층 이름 자체의 의미(예: "1층"/"1F"/"Level 1"이 같은 층을 뜻함)까지 함께 '
             '고려해 Gemini가 최종 매핑을 판단하도록 한다. IFC 파일쌍 하나당 API를 '
             '정확히 1회만 호출하며(층별로 반복 호출하지 않음), 결과는 이 세션 안에서 '
             '캐싱되어 화면이 새로고침돼도 재호출되지 않는다.',
    )
    if ai_floor_map_enabled:
        _secret_key = ''
        try:
            _secret_key = st.secrets.get('GOOGLE_API_KEY', '')
        except Exception:
            pass
        if _secret_key:
            google_api_key = _secret_key
            st.caption('✅ `.streamlit/secrets.toml`의 `GOOGLE_API_KEY`를 사용합니다.')
        else:
            google_api_key = st.text_input(
                'Google AI API Key', type='password',
                value=st.session_state.get('_google_api_key', ''),
                help='배포 시에는 이 입력창 대신 `.streamlit/secrets.toml`(로컬) 또는 '
                     'Streamlit Community Cloud의 App settings > Secrets(배포)에 '
                     '`GOOGLE_API_KEY = "..."`로 등록해두면 이 입력창 없이 자동으로 사용됩니다.',
            )
            st.session_state['_google_api_key'] = google_api_key
        llm_model_name = st.selectbox(
            '모델', options=['gemini-2.5-flash-lite', 'gemini-2.5-flash', 'gemini-3.1-flash-lite'],
            index=0,
            help='무료 할당량은 flash-lite 계열이 가장 넉넉한 경향입니다. 정확한 요청 한도(RPM/RPD)는 '
                 'Google이 공식 문서에 고정 수치를 게시하지 않으므로 https://aistudio.google.com/rate-limit '
                 '에서 프로젝트별로 직접 확인하시기 바랍니다.',
        )
        run_llm_floor_map = st.button('🚀 AI 층 매핑 실행 (API 1회 호출)', width='stretch')
        st.caption('⚠️ 이 버튼을 누를 때만 API가 호출됩니다. 체크박스만 켜는 것으로는 호출되지 않습니다.')
    else:
        google_api_key = None
        llm_model_name = None
        run_llm_floor_map = False


# ===================================================================
# 메인 UI
# ===================================================================

col_up1, col_up2 = st.columns(2)
with col_up1:
    file_a = st.file_uploader('전문가 IFC 업로드', type=['ifc'], key='upload_a')
with col_up2:
    file_b = st.file_uploader('AI 생성 IFC 업로드', type=['ifc'], key='upload_b')

if file_a and file_b:
    # 실제 파일 내용 기반 식별자 (다른 파일이 우연히 같은 층 이름을 가져도 캐시가 섞이지 않도록,
    # 아래 _build_plan_cached/_match_spaces_cached의 cache_tag에 사용)
    file_hash_a = hashlib.md5(file_a.getvalue()).hexdigest()[:10]
    file_hash_b = hashlib.md5(file_b.getvalue()).hexdigest()[:10]

    # 이전 실행에서 기록해둔 파일 해시와 다르면(=새 IFC로 교체됨) 캐시를 자동으로 비운다
    # (근본 조치: 정합성은 해시 기반 캐시 키로 이미 보장되지만, 이렇게 안 하면 안 쓰는
    # 이전 파일의 캐시 엔트리가 계속 쌓여 메모리를 차지하게 된다)
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

    # 층 자동매핑(고도 기준) - 트리의 배지 색상용. 계산이 가볍고(LLM 호출 없음) 결정론적이라
    # 별도 켜기/끄기 없이 항상 계산해서 참고용으로 보여준다.
    floor_mapping, floor_offset = fc.match_storeys(data_a['storeys'], data_b['storeys'])

    # AI(이름 문맥 + 표고) 층 매핑: 체크박스가 켜져 있고 API 키가 있을 때만, 그리고
    # (실행 버튼을 눌렀거나 이미 캐싱된 결과가 있을 때만) 계산한다. 파일쌍+모델명이
    # 캐시 키이므로, 같은 조합에 대해 재실행(rerun)은 API를 다시 부르지 않는다.
    llm_floor_mapping, llm_floor_detail = None, None
    if ai_floor_map_enabled:
        if not google_api_key:
            st.sidebar.warning('AI 층 매핑을 실행하려면 Google AI API Key를 입력하세요.')
        else:
            _llm_cache_tag = f'{file_hash_a}_{file_hash_b}'
            _llm_full_key = _SESSION_CACHE_PREFIX + f'llm_floor_{_llm_cache_tag}_{llm_model_name}'
            if run_llm_floor_map or _llm_full_key in st.session_state:
                # 버튼을 지금 막 눌러 실제로 API를 호출하는 경우에만 실시간 스트리밍 영역을 띄운다
                # (캐시 hit일 때는 호출 자체가 없으니 보여줄 스트림도 없다).
                _stream_area = None
                if run_llm_floor_map:
                    st.markdown('#### 🔎 Gemini 실시간 추론 진행')
                    _stream_area = st.empty()
                    _stream_area.code('(응답 대기 중...)', language='json')
                try:
                    llm_floor_mapping, llm_floor_detail = _match_storeys_llm_cached(
                        data_a['storeys'], data_b['storeys'], floor_offset, floor_mapping,
                        google_api_key, llm_model_name, _llm_cache_tag,
                        force_recompute=run_llm_floor_map, stream_placeholder=_stream_area,
                    )
                    if _stream_area is not None:
                        st.caption('✅ 위 스트리밍 원본 응답을 파싱한 결과가 아래 배지/상세근거에 반영되었습니다.')
                except lsm.LlmStoreyMatchError as e:
                    st.sidebar.error(f'AI 층 매핑 실패: {e}')

    # 배지 표시는 AI 매핑 결과가 있으면 그것을 우선 사용하고, 없으면 표고 기반 매핑을 사용
    _badge_source = llm_floor_mapping if llm_floor_mapping is not None else floor_mapping
    _badge_source_inv = _invert_mapping(_badge_source)
    floor_badges_a, floor_badges_b = _floor_pair_badges(_badge_source)

    with st.sidebar:
        st.divider()
        st.header('🏢 층 선택')
        floor_auto_follow = st.checkbox(
            '매핑된 층 자동 함께 선택', value=True,
            help='한쪽에서 층을 체크하면, 배지가 같은 색인(매핑된) 반대편 층도 자동으로 '
                 '함께 선택됩니다. 대응되는 층이 없는(⬜) 경우에는 반대편은 그대로 둡니다.',
        )
        if llm_floor_mapping is not None:
            st.caption(
                '같은 색 배지 = AI(이름 문맥 + 표고 유사도)가 매핑한 층입니다. '
                '⬜는 대응되는 층을 못 찾은 경우입니다.'
            )
            with st.expander('AI 매핑 상세 근거 보기'):
                for d in llm_floor_detail:
                    st.caption(f"`{d['a_name']}` → `{d['b_name']}` (확신도: {d['confidence']}) — {d['reason']}")
        else:
            st.caption(
                '같은 색 배지가 붙은 층끼리 고도 기준 자동 매핑된 층입니다 '
                f'(오프셋 {floor_offset:.0f}mm 보정). ⬜는 대응되는 층을 못 찾은 경우입니다.'
            )
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

    # 층 선택이 바뀌면 이전 선택된 공간 정보는 초기화
    _cur_key = (selected_a_name, selected_b_name)
    if st.session_state.get('_last_storey_pair') != _cur_key:
        st.session_state.pop('left_selected_guid', None)
        st.session_state.pop('right_selected_guid', None)
        st.session_state.pop('left_space_dropdown', None)
        st.session_state.pop('right_space_dropdown', None)
        st.session_state['_last_storey_pair'] = _cur_key

    st.markdown(f"### 비교 중: 전문가 `{selected_a_name}` ↔ AI `{selected_b_name}`")
    st.caption('층은 왼쪽 사이드바 "🏢 층 선택"에서 바꿀 수 있습니다.')

    plan_a = _build_plan_cached(data_a['storeys'], selected_a_name, f'left_{file_hash_a}')
    plan_b = _build_plan_cached(data_b['storeys'], selected_b_name, f'right_{file_hash_b}')

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

    col_left, col_right = st.columns(2)
    with col_left:
        detail_left, new_left = _render_plot_and_get_detail(
            '전문가 IFC', data_a, selected_a_name, plan_a, 'left', pair_labels=pair_labels_a)
    with col_right:
        detail_right, new_right = _render_plot_and_get_detail(
            'AI IFC', data_b, selected_b_name, plan_b, 'right', pair_labels=pair_labels_b)

    if detail_left is not None or detail_right is not None:
        _render_legend()

    changed = False
    if new_left:
        st.session_state['left_selected_guid'] = new_left
        changed = True
        if auto_map_enabled and new_left in space_a_to_b:
            st.session_state['right_selected_guid'] = space_a_to_b[new_left]
    if new_right:
        st.session_state['right_selected_guid'] = new_right
        changed = True
        if auto_map_enabled and new_right in space_b_to_a:
            st.session_state['left_selected_guid'] = space_b_to_a[new_right]
    if changed:
        st.rerun()  # 하이라이트/자동매핑 반영을 위해 갱신된 session_state로 즉시 재실행

    _render_comparison_tables(detail_left, detail_right, '전문가', 'AI')
else:
    st.info('좌측/우측에 전문가 IFC와 AI 생성 IFC를 각각 업로드해주세요.')
