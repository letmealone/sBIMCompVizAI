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
from shapely.geometry import Polygon, Point, box
from shapely.ops import unary_union

import ifc_to_excel as ite  # 내외벽 판정(_determine_wall_classification), 면적계산(_area_columns) 등 재사용

_SETTINGS = geom.settings()
_SETTINGS.set('use-world-coords', True)

# 평면도에 그릴 대상 클래스. Space는 클릭 가능(색상 채움), 나머지는 참고용 윤곽선만 표시.
PLAN_STRUCTURAL_CLASSES = (
    'IfcWall', 'IfcWallStandardCase', 'IfcColumn', 'IfcBeam',
    'IfcSlab', 'IfcCurtainWall', 'IfcDoor', 'IfcWindow',
)


# ===================================================================
# 1. IFC 로딩 및 층 매핑
# ===================================================================

def load_ifc(path):
    """IFC 파일을 열고 앱에서 바로 쓸 수 있는 형태로 구조화."""
    ifc_file = ifcopenshell.open(path)

    storeys = []
    for s in ifc_file.by_type('IfcBuildingStorey'):
        storeys.append({'Name': s.Name, 'Elevation': s.Elevation, 'entity': s})
    storeys.sort(key=lambda x: (x['Elevation'] is None, x['Elevation']))

    wall_classification = ite._determine_wall_classification(ifc_file)
    # 벽 외 모든 구조부재(슬래브/기둥/보/문/창/커튼월 등) 대상 범용 내/외부 판정.
    # 비교 엑셀(comparison_export.py)에서 사용 - 여기서 미리 계산해 캐싱하면
    # 공간별로 반복 호출해도 매번 다시 계산하지 않는다.
    element_classification = ite._determine_element_classification(ifc_file)

    return {
        'ifc_file': ifc_file,
        'storeys': storeys,
        'wall_classification': wall_classification,
        'element_classification': element_classification,
    }


def match_storeys(storeys_a, storeys_b, gap_cost=1000.0):
    """고도(Elevation) 기반 층 매핑 (ifc_compare_core.py의 build_floor_mapping과 동일한 방식,
    이 앱은 openai/httpx 등 무거운 의존성 없이 독립 실행되도록 여기서 간단히 재구현).
    반환: A_Name -> B_Name 매핑 dict, 오프셋(mm)."""
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


# ===================================================================
# 1b. 공간(Space) 자동 매핑 (면적 + centroid 좌표 오차 기준)
# ===================================================================
# 주의(실측으로 확인한 사실): 서로 다른 저작자가 만든 두 IFC는 좌표계 원점이 다를 수 있어
# (실제로 샘플 파일 쌍에서 그랬음) centroid 좌표를 그냥 비교하면 항상 실패한다.
# 반면 면적은 저작자와 무관하게 거의 그대로 보존되므로, 면적이 비슷한 후보쌍들에서
# "좌표계 평행이동 오프셋"을 먼저 통계적으로 추정한 뒤 그 오프셋을 보정해 centroid 거리를
# 비교한다. 회전 차이는 다루지 않는다(검증한 샘플은 회전 없이 평행이동만 달랐음 - 다른
# 모델 쌍은 회전까지 다를 경우 이 방식이 통하지 않을 수 있음).

def _offset_candidates(spaces_a, spaces_b, area_thresh):
    """면적차가 area_thresh 이내인 모든 (A,B) 후보쌍의 offset(=centroid_B - centroid_A) 벡터."""
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
    """후보 offset들 중 가장 밀집된 클러스터의 평균을 좌표계 오프셋으로 추정.
    (정답 매칭들은 전부 같은 오프셋에 모이고, 우연히 면적만 비슷한 오탐은 흩어지므로
    가장 큰 클러스터가 진짜 오프셋일 가능성이 높다)"""
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
    """면적(㎡) 오차와 좌표계 오프셋 보정 후 centroid 거리(m) 오차가 각각 임계값 이내인
    공간을 1:1 그리디 매칭(가까운 거리 우선).
    반환: (a_to_b, b_to_a, offset, match_info)
      - a_to_b/b_to_a: {GlobalId: GlobalId} 매핑 dict
      - offset: 추정된 좌표계 평행이동 (dx, dy) 또는 매칭 후보가 없으면 None
      - match_info: 매칭된 쌍의 상세 리스트(diagnostic/표시용)
    """
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


# ===================================================================
# 2. 지오메트리 (실제 footprint 폴리곤 추출)
# ===================================================================

def get_footprint_polygon(ent, tol=0.05):
    """엔티티의 바닥면(최저 Z 근처) 삼각형들을 shapely로 합쳐 실제 footprint 폴리곤 반환.
    형상이 없거나 계산 실패시 None. tol: 바닥면으로 간주할 Z 허용오차(m)."""
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
    """폴리곤 내부를 spacing(m) 간격 격자로 채운 점 목록. Plotly 클릭 히트영역용
    (폴리곤 전체를 '클릭 가능 영역'으로 만들기 위해 보이지 않는 마커를 촘촘히 깔아둔다)."""
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
    """shapely (Multi)Polygon -> Plotly에 그릴 (x리스트, y리스트) 반환.
    여러 폴리곤/구멍은 None으로 구분해 하나의 트레이스에 이어붙인다."""
    xs, ys = [], []

    def _add_ring(coords):
        cx, cy = zip(*coords)
        xs.extend(cx); xs.append(None)
        ys.extend(cy); ys.append(None)

    geoms = poly.geoms if poly.geom_type == 'MultiPolygon' else [poly]
    for g in geoms:
        _add_ring(list(g.exterior.coords))
    return xs, ys


# ===================================================================
# 3. 층별 요소 수집
# ===================================================================

def get_elements_for_storey(storey_entity, classes=None):
    """해당 층에 속한 요소 목록.
    storey_entity: load_ifc()가 반환한 storeys 리스트의 원소(dict, 'entity' 키에 실제 IfcBuildingStorey).
    구조부재(벽/기둥 등)는 IfcRelContainedInSpatialStructure(ContainsElements)로,
    IfcSpace는 대개 IfcRelAggregates(IsDecomposedBy)로 층에 연결되므로 둘 다 확인한다."""
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
    """해당 층 IfcSpace들의 이름 구성을 경량으로 요약한다.
    get_footprint_polygon() 등 지오메트리 계산(삼각분할)을 전혀 하지 않고 관계 순회만
    하므로 매우 가볍다 - AI 층 매핑 프롬프트에서 '공간 구성 유사도' 신호로 쓰기 위한 것.
    반환: {'count': 전체 공간 수, 'names': 이름 리스트(등장 순, 중복 포함, 최대 max_names개)}
    """
    spaces = get_elements_for_storey(storey_entity, classes={'IfcSpace'})
    names = [(sp.LongName or sp.Name or '(이름없음)') for sp in spaces]
    return {'count': len(names), 'names': names[:max_names]}


def get_element_hover_info(ent, wall_classification=None):
    """평면도에서 부재 위에 마우스를 올렸을 때 보여줄 요약정보(HTML, <br> 개행) 생성.
    치수/면적/재질 추출은 ifc_to_excel.py에 이미 있는 로직(좌표 기반 직접계산, Pset 폴백
    포함)을 그대로 재사용해 엑셀 추출 결과와 수치가 항상 일치하도록 한다(중복 구현 방지).
    wall_classification: ifc_to_excel._determine_wall_classification() 반환값. 주어지고
    이 부재가 벽이면 내/외벽 판정 결과도 함께 표시한다."""
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
    """해당 층의 Space + 구조요소들의 footprint 폴리곤을 미리 계산해 리스트로 반환.
    반환: {'spaces': [{'guid','name','polygon'}...],
           'structural': [{'guid','class','name','polygon','hover'}...]}
    (지오메트리 계산은 비용이 있으므로 앱에서 층 변경시에만 1회 호출하도록 캐싱 권장.
    hover 텍스트도 여기서 미리 만들어 캐싱 대상에 포함시킨다 - 매 렌더링마다 다시
    치수/속성을 파싱하면 비용이 반복되므로 층 변경시 1회만 계산되게 하기 위함)"""
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


# ===================================================================
# 4. 클릭된 Space의 상세 정보
# ===================================================================

EQUIPMENT_CLASSES = ('IfcLightFixture', 'IfcSensor', 'IfcFireSuppressionTerminal', 'IfcAlarm')


def get_space_related_elements(ifc_file, space_entity):
    """해당 Space와 RelSpaceBoundary로 연결된 부재 목록 (벽/기둥/문/창/바닥 등 경계형성 요소).
    GlobalId 기준으로 중복 제거한다: 같은 부재(예: 기둥 하나)가 하나의 Space와 여러 개의
    별도 경계면(RelSpaceBoundary 레코드)으로 연결되는 경우가 실제로 있음을 확인했다
    (예: 샘플 파일 Space-S-01에 접한 기둥은 물리적으로 7개인데, RelSpaceBoundary 레코드는
    23건 - 기둥 하나당 여러 면이 각각 별도 레코드로 잡힘). 개수 집계는 물리적 개체 수
    기준이어야 하므로 여기서 dedup한다."""
    by_guid = {}
    for rel in ifc_file.by_type('IfcRelSpaceBoundary'):
        if rel.RelatingSpace == space_entity and rel.RelatedBuildingElement is not None:
            el = rel.RelatedBuildingElement
            by_guid[el.GlobalId] = el
    return list(by_guid.values())


def get_space_contained_equipment(ifc_file, space_entity, classes=EQUIPMENT_CLASSES):
    """해당 Space '안에' 배치된 설비(조명/센서/소방장치 등) 목록.
    이런 설비는 RelSpaceBoundary(경계형성)가 아니라 IfcRelContainedInSpatialStructure
    (공간적 포함관계)로 Space와 연결된다 (실측으로 확인한 IFC 구조)."""
    equipment = []
    for rel in ifc_file.by_type('IfcRelContainedInSpatialStructure'):
        if rel.RelatingStructure == space_entity:
            for el in rel.RelatedElements:
                if el.is_a() in classes:
                    equipment.append(el)
    return equipment


def _wall_display_category(result):
    """내/외벽 판정을 표시용 카테고리로 변환 (4분류):
    - '내벽' -> '내부' (1차/2차로 확정)
    - '내벽(추정-관계기반)' -> '내부(추정)' (3차, ConnectionGeometry 없이 관계 개수만으로 추정 - 확정보다 약함)
    - '판정불가'(1차/2차/3차 모두 근거 없음) -> '외부(판정불가)'
    - 그 외('외벽'=1차 확정, '외벽(추정)'=2차 확정) -> '외부(판정됨)'
    즉 외벽/내벽 모두 '확정'과 '추정/근거없음'을 구분해서 보여준다."""
    if result == '내벽':
        return '내부', '내벽'
    if result == '내벽(추정-관계기반)':
        return '내부(추정)', result
    if result == '판정불가':
        return '외부(판정불가)', '외벽(판정불가)'
    return '외부(판정됨)', result  # '외벽' / '외벽(추정)'


_AREA_APPORTION_CLASSES = ('IfcSlab', 'IfcRoof', 'IfcCovering')
_APPORTION_BUFFER_M = 0.3  # 공간 폴리곤을 이만큼(m) 부풀려서 부재 footprint와 겹치는 부분을 봄.
# 실측으로 튜닝한 값: 한 공간에만 속한 짧은 벽/바닥은 이 정도면 거의 100% 잡히고,
# 여러 공간에 걸친 긴 벽/바닥은 실제로 걸친 비율만큼만 낮게 나옴 (예: 방 하나에 딸린 벽은
# 비율 1.0, 여러 방을 관통하는 긴 벽은 0.15~0.25 수준으로 나뉘어 원래 문제를 해결함).

# 면적이 물리적으로 "의미있는" 지표인 클래스만 여기 포함시킨다.
#   - Slab/Roof/Covering: 압출단면(=평면 footprint)으로 안정적으로 모델링되어 신뢰할 만함
#     (여러 공간에 걸칠 수 있어 안분 적용).
#   - Door/Window/CurtainWall: ite._area_columns()가 "시스템 전체 bounding치수(폭x높이,
#     두께 제외)" 기반으로 계산하도록 수정되어(부품/멀리언 단위 메쉬 합산이 아님) 이제
#     신뢰할 만함(안분 불필요 - 보통 하나의 특정 공간에 귀속되는 개구부이므로 전체값 사용).
_STRUCTURAL_AREA_MEANINGFUL_CLASSES = {'IfcSlab', 'IfcRoof', 'IfcCovering', 'IfcDoor', 'IfcWindow', 'IfcCurtainWall'}
# (참고) 기둥/보/부재/계단/문/창/커튼월/난간은 위 화이트리스트에 없어 자동으로
# "면적 없음(개수만)" 처리된다 - 위 주석 참고.


_boundary_index_cache = {}  # id(ifc_file) -> {element_guid: [IfcRelSpaceBoundary, ...]}
# ifc_file 객체 하나당 한 번만 전체 RelSpaceBoundary를 훑어 인덱싱해두고 재사용한다.
# (부재별로 매번 ifc_file.by_type('IfcRelSpaceBoundary') 전체를 다시 훑으면, 공간이
# 많은 파일을 일괄 비교할 때 매우 느려지므로 이 캐시가 필수적이다.)


def _get_boundary_index(ifc_file):
    key = id(ifc_file)
    if key not in _boundary_index_cache:
        index = defaultdict(list)
        for r in ifc_file.by_type('IfcRelSpaceBoundary'):
            el = r.RelatedBuildingElement
            if el is not None:
                index[el.GlobalId].append(r)
        _boundary_index_cache[key] = index
    return _boundary_index_cache[key]


def _member_axis_segments(member_entity):
    """부재(벽 등)의 Axis(축) 표현에서 로컬좌표계 기준 폴리선을 구간(세그먼트) 리스트로
    추출한다. 직선벽이면 세그먼트 1개, 꺾인(폴리라인) 벽이면 여러 개.
    반환: [((x1,y1),(x2,y2)), ...] (부재 로컬좌표, mm 단위, 원본단위 그대로) 또는
    None(Axis 표현 자체가 없는 경우 - Tessellation만 있는 일부 부재 등)."""
    rep = getattr(member_entity, 'Representation', None)
    if rep is None:
        return None
    axis_pts = None
    for r in rep.Representations:
        if r.RepresentationIdentifier != 'Axis':
            continue
        for item in r.Items:
            if item.is_a('IfcIndexedPolyCurve'):
                pts = item.Points
                if pts.is_a('IfcCartesianPointList3D'):
                    axis_pts = [(p[0], p[1]) for p in pts.CoordList]
                elif pts.is_a('IfcCartesianPointList2D'):
                    axis_pts = [(p[0], p[1]) for p in pts.CoordList]
            elif item.is_a('IfcPolyline'):
                axis_pts = [(p.Coordinates[0], p.Coordinates[1]) for p in item.Points]
            if axis_pts:
                break
        if axis_pts:
            break
    if not axis_pts or len(axis_pts) < 2:
        return None
    return [(axis_pts[i], axis_pts[i + 1]) for i in range(len(axis_pts) - 1)]


def _nearest_segment_index(local_xy, segments):
    """local_xy(부재 로컬좌표, mm)에 가장 가까운 축 세그먼트의 인덱스를 반환
    (점을 각 세그먼트 선분에 투영해 클램프한 최단거리 기준)."""
    pt = np.array(local_xy)
    best_idx, best_dist = 0, None
    for i, (p1, p2) in enumerate(segments):
        a, b = np.array(p1), np.array(p2)
        ab = b - a
        denom = np.dot(ab, ab)
        t = np.dot(pt - a, ab) / denom if denom > 1e-9 else 0.0
        t = max(0.0, min(1.0, t))
        proj = a + t * ab
        dist = np.linalg.norm(pt - proj)
        if best_dist is None or dist < best_dist:
            best_dist, best_idx = dist, i
    return best_idx


def _segment_side(local_xy, seg_p1, seg_p2):
    """축 세그먼트 하나를 기준으로 한 국소좌표계에서, local_xy가 세그먼트의 어느 편측
    (+/-, 진행방향 기준 왼쪽/오른쪽)에 있는지 판정한다."""
    a, b = np.array(seg_p1), np.array(seg_p2)
    direction = b - a
    length = np.linalg.norm(direction)
    if length < 1e-6:
        return None
    x_axis = direction / length
    y_axis = np.array([-x_axis[1], x_axis[0]])  # 평면상 90도 회전 = 두께 방향
    rel = np.array(local_xy) - a
    return '+' if np.dot(rel, y_axis) >= 0 else '-'


def _space_side_of_member(space_entity, member_m_inv, axis_segments=None):
    """공간의 footprint 중심점을 부재(벽 등)의 로컬 좌표계로 변환해 편측('+'/'-')을 판정한다.

    axis_segments가 주어지면(부재의 Axis가 여러 구간으로 꺾인 경우), 공간 중심점에서
    가장 가까운 축 세그먼트를 찾아 '그 세그먼트만의 국소좌표계' 기준으로 편측을 판정한다
    - 꺾인(L자형 등) 벽 하나를 부재 전체의 단일 좌표계로만 판단하면, 꺾인 두 구간이
    실제로는 서로 다른 물리적 면인데도 부재 로컬좌표계 하나로는 구분 못하는 경우가
    실측으로 확인되어(구간별 국소좌표계 미적용시 정합성 통과 147개 -> 이 구간분리
    적용 후 추가 개선, 아래 docstring 참고) 이 방식을 추가했다.
    axis_segments가 없으면(직선 벽이거나 Axis 표현이 없는 부재) 부재 전체 로컬좌표계
    기준(기존 방식)으로 판정한다.

    단위 주의(실측으로 확인된 함정): ifcopenshell.util.placement는 IFC 원본단위(보통
    mm) 그대로 반환하는데, get_footprint_polygon()은 미터 단위로 반환한다. 이 둘을
    그대로 섞어 곱하면 단위가 1000배 어긋나 조용히 틀린(값이 다 뭉개지는) 결과가
    나오므로, 여기서 반드시 1000배 해서 mm로 맞춘다."""
    poly = get_footprint_polygon(space_entity)
    if poly is None or poly.is_empty:
        return None
    cx, cy = poly.centroid.x * 1000.0, poly.centroid.y * 1000.0
    try:
        import ifcopenshell.util.placement as plc
        space_m = plc.get_local_placement(space_entity.ObjectPlacement)
    except Exception:
        return None
    cz = space_m[2, 3]  # 이미 mm(원본단위) - get_local_placement 반환값이라 그대로 사용
    local_pt = member_m_inv @ np.array([cx, cy, cz, 1.0])

    if axis_segments:
        idx = _nearest_segment_index((local_pt[0], local_pt[1]), axis_segments)
        p1, p2 = axis_segments[idx]
        side = _segment_side((local_pt[0], local_pt[1]), p1, p2)
        if side is not None:
            return side
        # 세그먼트 방향이 퇴화(길이 0)된 예외적 경우만 아래 전체좌표계 방식으로 폴백

    return '+' if local_pt[1] >= 0 else '-'


def _relspaceboundary_precise_areas(ifc_file, member_entity, own_side_area_m2, tol_factor=1.10):
    """RelSpaceBoundary.ConnectionGeometry 기반으로 member_entity가 접한 각 Space별
    정밀 귀속 면적을 계산해보고, 정합성 검증까지 통과하면 반환한다.

    설계(사용자 요청): "부재 전체면적 = RelSpaceBoundary 면적 합 + 나머지(어떤 공간과도
    접하지 않은 부분)"라는 논리를 1차로 적용해보되, 면(앞/뒤) 구분이 필요하다.

    면 구분 방식(개선 이력): 처음에는 "공간의 배치행렬 기준 경계면 법선벡터"로 같은면/
    반대면을 나눴으나, 실측 결과 서로 마주보는 인접한 두 공간(예: 벽 하나를 사이에 둔
    Space-P-02/Space-P-05)이 반대 면인데도 같은 방향 법선으로 계산되는 사례를 발견했다
    (공간마다 로컬 좌표계 회전 관례가 일관되지 않아 발생하는 것으로 추정). 그래서
    "부재(벽) 자신의 로컬 좌표계 기준으로, 각 공간 footprint 중심점이 어느 쪽(+Y/-Y)에
    있는가"로 방식을 바꿨다 - 부재는 좌표계가 하나뿐이라 기준이 일관된다. 실측으로
    개선 확인: 같은 건물 전체 벽 기준 정합성 통과 110개 -> 147개로 증가, 실패 149->112.

    그래도 "어느 편측 그룹의 합계든 own_side_area_m2(부재 한쪽 면 전체 면적, 예: 벽의
    Gross_Side_Area)의 tol_factor배를 넘으면(물리적으로 불가능)" 정합성 실패로 보고
    None을 반환하는 안전장치는 그대로 둔다 - 이때 호출부(_apportioned_area)가 기존
    footprint 버퍼 근사 방식으로 폴백한다. (여전히 112개는 이 개선으로도 못 걸러지는
    더 복잡한 사례이거나 진짜 데이터 이상치일 수 있음 - 안전하게 폴백됨)

    반환: {space_guid: 면적(㎡)} (정합성 통과시, 같은 공간이 여러 조각으로 나뉘면 합산됨)
          또는 None (ConnectionGeometry 데이터 없음/정합성 실패).
    """
    import ifcopenshell.util.placement as plc

    try:
        member_m = plc.get_local_placement(member_entity.ObjectPlacement)
        member_m_inv = np.linalg.inv(member_m)
    except Exception:
        return None

    axis_segments = _member_axis_segments(member_entity)  # 꺾인 벽이면 구간별 판정에 사용

    rels = _get_boundary_index(ifc_file).get(member_entity.GlobalId, [])
    boundaries = []
    space_side_cache = {}
    for r in rels:
        space = r.RelatingSpace
        cg = r.ConnectionGeometry
        if space is None or cg is None:
            continue
        surf = cg.SurfaceOnRelatingElement
        if surf is None or not surf.is_a('IfcCurveBoundedPlane'):
            continue
        pts = ite._polyline_points_2d(surf.OuterBoundary)
        area_mm2 = ite._shoelace_area(pts) if pts else None
        if area_mm2 is None:
            continue

        if space.GlobalId not in space_side_cache:
            space_side_cache[space.GlobalId] = _space_side_of_member(space, member_m_inv, axis_segments)
        side = space_side_cache[space.GlobalId]
        if side is None:
            continue
        boundaries.append({'space_guid': space.GlobalId, 'area_m2': area_mm2 / 1e6, 'side': side})

    if not boundaries or not own_side_area_m2 or own_side_area_m2 <= 0:
        return None

    group_sums = defaultdict(float)
    for b in boundaries:
        group_sums[b['side']] += b['area_m2']

    limit = own_side_area_m2 * tol_factor
    if any(s > limit for s in group_sums.values()):
        return None  # 정합성 실패 -> 호출부가 근사 방식으로 폴백

    per_space = defaultdict(float)
    for b in boundaries:
        per_space[b['space_guid']] += b['area_m2']
    return dict(per_space)



def _apportioned_area(ifc_file, member_entity, target_space_entity, own_side_area_m2, space_footprint,
                       tol_factor=1.10):
    """공간 하나에 귀속되는 부재 면적을 계산한다 (1차 정밀, 실패시 2차 근사로 자동 폴백).
    1차: RelSpaceBoundary 정밀 계산 - 정합성 검증까지 통과하면 이 값을 그대로 쓴다.
    2차(정밀 계산 데이터 없음/정합성 실패시): 기존 footprint 버퍼 근사.
    반환: (면적, 산출방식 라벨) - 라벨은 진단/투명성 목적으로 호출부가 필요시 노출 가능."""
    precise = _relspaceboundary_precise_areas(ifc_file, member_entity, own_side_area_m2, tol_factor)
    if precise is not None and target_space_entity.GlobalId in precise:
        return precise[target_space_entity.GlobalId], '정밀(RelSpaceBoundary)'

    fraction = _space_portion_fraction(member_entity, space_footprint)
    if fraction is None:
        return own_side_area_m2, '실패-전체값사용(과다산정 가능)'
    return own_side_area_m2 * fraction, '근사(footprint버퍼)'


def _space_portion_fraction(member_entity, space_footprint, buffer_dist=_APPORTION_BUFFER_M):
    """member_entity(벽/바닥/지붕/천장 등)의 footprint 중, space_footprint에 buffer_dist(m)만큼
    부풀린 영역과 겹치는 비율(0~1)을 반환. 부재가 여러 공간에 걸쳐 있을 때 '이 공간에 해당하는
    부분'만 면적에 반영하기 위한 안분 비율이다. 계산 실패(footprint 없음 등)시 None 반환
    (호출부에서 안분 없이 전체 면적을 쓰도록 폴백)."""
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


def get_space_wall_segment_polygon(ifc_file, wall_entity, space_entity, wall_footprint_polygon):
    """평면도 표시용: 벽 전체가 아니라 '선택된 공간에 실제로 접한 부분만' 잘라낸 폴리곤을
    반환한다. RelSpaceBoundary.ConnectionGeometry(경계면 폴리곤)를 이용해, 그 경계가
    벽의 길이방향 축 기준 어느 구간([x_min,x_max])에 해당하는지 구한 뒤, 벽의 전체
    footprint를 그 구간만큼만 클리핑한다.

    계산 절차(벽 자신의 로컬좌표계 기준 - _space_side_of_member와 같은 방식):
      1. 이 벽-공간 쌍의 RelSpaceBoundary들에서 경계 폴리곤 점들을 월드좌표로 변환
      2. 벽의 로컬좌표계로 재변환해 길이방향(X) 범위[x_min,x_max] 추출
      3. 벽의 전체 footprint(월드,m)를 벽 로컬좌표(mm)로 옮겨 그 X범위로 클리핑
      4. 다시 월드좌표(m)로 되돌려 반환

    반환: 클리핑된 Polygon, 또는 계산 불가시(ConnectionGeometry 없음 등) None
    (호출부는 None이면 벽 전체 폴리곤을 그대로 표시하는 폴백을 쓰면 된다)."""
    import ifcopenshell.util.placement as plc

    try:
        member_m = plc.get_local_placement(wall_entity.ObjectPlacement)
        member_m_inv = np.linalg.inv(member_m)
    except Exception:
        return None

    rels = _get_boundary_index(ifc_file).get(wall_entity.GlobalId, [])
    local_xs = []
    for r in rels:
        if r.RelatingSpace != space_entity:
            continue
        cg = r.ConnectionGeometry
        if cg is None:
            continue
        surf = cg.SurfaceOnRelatingElement
        if surf is None or not surf.is_a('IfcCurveBoundedPlane'):
            continue
        pts2d = ite._polyline_points_2d(surf.OuterBoundary)
        if not pts2d:
            continue
        try:
            plane_m = plc.get_axis2placement(surf.BasisSurface.Position)
            space_m = plc.get_local_placement(space_entity.ObjectPlacement)
        except Exception:
            continue
        world_m = space_m @ plane_m
        for (u, v) in pts2d:
            world_pt = world_m @ np.array([u, v, 0.0, 1.0])
            local_pt = member_m_inv @ np.array([world_pt[0], world_pt[1], world_pt[2], 1.0])
            local_xs.append(local_pt[0])

    if not local_xs:
        return None  # ConnectionGeometry 데이터 없음 -> 호출부가 벽 전체로 폴백

    margin = 50.0  # mm, 경계에 딱 붙어 잘리는 것을 막기 위한 여유
    x_min, x_max = min(local_xs) - margin, max(local_xs) + margin

    if wall_footprint_polygon is None or wall_footprint_polygon.is_empty:
        return None
    coords_local = []
    for (wx, wy) in wall_footprint_polygon.exterior.coords:
        local_pt = member_m_inv @ np.array([wx * 1000.0, wy * 1000.0, 0.0, 1.0])
        coords_local.append((local_pt[0], local_pt[1]))
    local_poly = Polygon(coords_local)
    if not local_poly.is_valid or local_poly.is_empty:
        return None

    minx, miny, maxx, maxy = local_poly.bounds
    clip_box = box(x_min, miny - margin, x_max, maxy + margin)
    clipped_local = local_poly.intersection(clip_box)
    if clipped_local.is_empty:
        return None

    def _to_world_xy(pt):
        world_pt = member_m @ np.array([pt[0], pt[1], 0.0, 1.0])
        return (world_pt[0] / 1000.0, world_pt[1] / 1000.0)

    if clipped_local.geom_type == 'Polygon':
        target = clipped_local
    elif clipped_local.geom_type == 'MultiPolygon':
        target = max(clipped_local.geoms, key=lambda g: g.area)
    else:
        return None
    world_coords = [_to_world_xy(p) for p in target.exterior.coords]
    result = Polygon(world_coords)
    return result if result.is_valid and not result.is_empty else None


def build_space_detail(ifc_file, wall_classification, space_entity):
    """클릭된 Space 1개에 대한 요약 정보:
    - 접한 구조재 개수(전체) + 벽 내/외부 구분(좌우대칭 이진 + 상세근거 병기)
    - 관련된 모든 부재 유형별 합산 면적(계산 가능한 모든 클래스, 벽은 별도 처리)
    - 공간 내 설비(조명/센서/소방장치/경보기) 개수
    """
    related = get_space_related_elements(ifc_file, space_entity)
    equipment = get_space_contained_equipment(ifc_file, space_entity)
    space_footprint = get_footprint_polygon(space_entity)  # 안분 계산 기준(이 공간의 실제 바닥형상)

    # 평면도 표시용: 벽마다 '이 공간에 실제로 접한 부분만' 잘라낸 폴리곤(계산 가능한 것만).
    # None이면 호출부(build_plan_figure)가 벽 전체를 표시하는 것으로 폴백한다.
    wall_segment_polygons = {}
    for e in related:
        if not e.is_a('IfcWall'):
            continue
        wall_footprint = get_footprint_polygon(e)
        wall_segment_polygons[e.GlobalId] = get_space_wall_segment_polygon(
            ifc_file, e, space_entity, wall_footprint)

    # 진단용: 이 공간에 접한 벽 중 몇 개가 정밀표시(ConnectionGeometry 기반 클리핑)되고
    # 몇 개가 폴백(벽 전체 표시)됐는지 - 화면만 봐서는 구분이 안 되므로 앱에서 캡션으로
    # 보여줄 수 있게 여기서 집계해둔다.
    wall_segment_stats = {
        'precise': sum(1 for v in wall_segment_polygons.values() if v is not None),
        'fallback': sum(1 for v in wall_segment_polygons.values() if v is None),
    }

    class_counts = Counter(e.is_a() for e in related)

    # 벽 내/외부 구분 (좌우 대칭 이진) + 상세 판정(원래 4분류) 병기
    # 면적은 벽 전체 값이 아니라 이 공간에 걸친 부분만 안분해서 합산한다 (여러 공간에 이어진
    # 벽의 전체 길이가 그대로 잡히던 문제를 해결 - _space_portion_fraction 참고).
    wall_simple_counts = Counter()
    wall_simple_area = Counter()
    wall_detail_counts = Counter()
    wall_area_apportioned = True  # 하나라도 안분 실패(폴백)하면 False로 내려 라벨에 표시
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
            area_val, method = _apportioned_area(ifc_file, e, space_entity, v, space_footprint)
            if method == '실패-전체값사용(과다산정 가능)':
                wall_area_apportioned = False
            wall_simple_area[simple] += area_val

    # 벽 이외 관련 부재: 계산 가능한 모든 클래스에 대해 면적 산정 시도 ("가능한 경우"만 채워짐)
    # 바닥/지붕/천장(IfcSlab/IfcRoof/IfcCovering)은 벽과 마찬가지로 여러 공간에 걸칠 수 있어
    # footprint 안분을 적용한다. 기둥/문/창 등은 부재가 길이 방향으로 여러 공간에 나뉘는
    # 개념이 아니라(양쪽 공간이 공유하는 고정된 개구부/단면) 전체 값을 그대로 쓴다.
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
                if apportioned:
                    area_val, method = _apportioned_area(ifc_file, e, space_entity, cols['면적(㎡)'], space_footprint)
                    if method == '실패-전체값사용(과다산정 가능)':
                        any_fallback = True
                else:
                    area_val = cols['면적(㎡)']
                total += area_val
                n_ok += 1
        note = ''
        if apportioned:
            note = '(공간 귀속분 안분)' if not any_fallback else '(일부 안분실패-전체값 폴백)'
        area_by_class[cls] = {
            '면적합계(㎡)': round(total, 2) if n_ok else None,
            '산출가능/전체': f'{n_ok}/{len(ents)}',
            '비고': note,
        }

    # 설비 개수 (구조재와 별도 집계)
    equipment_counts = Counter(e.is_a() for e in equipment)

    # Space 자신의 면적
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
        # 평면도 하이라이트용: 각 관련 부재 GlobalId -> 표시 카테고리
        'highlight_map': _build_highlight_map(related, equipment, wall_classification),
        'wall_segment_polygons': wall_segment_polygons,
        'wall_segment_stats': wall_segment_stats,
    }


def build_space_structural_breakdown(ifc_file, element_classification, space_entity):
    """클릭된 Space 1개에 대해, 벽 뿐 아니라 '모든 구조부재 클래스'를 대상으로
    (a) 내/외부 구분이 있는 집계와 (b) 구분 없는 집계를 각각 만든다 - 엑셀 비교 시트용.
    (build_space_detail은 벽 전용 세부정보이고, 이 함수는 그걸 모든 클래스로 일반화한 것 -
    설비(EQUIPMENT_CLASSES)는 여기 포함하지 않는다 - get_space_related_elements 자체가
    RelSpaceBoundary 기반이라 설비(RelContainedInSpatialStructure 기반)는 애초에 안 잡힌다.)

    면적 산정 규칙 (build_space_detail과 동일한 안분 로직 재사용, 화이트리스트 방식 -
    검증되지 않은 클래스는 기본적으로 면적을 계산하지 않고 개수만 집계한다):
      - IfcWall/IfcWallStandardCase: Qto_WallBaseQuantities.Gross_Side_Area, footprint 안분.
      - _STRUCTURAL_AREA_MEANINGFUL_CLASSES(Slab/Roof/Covering, Door/Window/CurtainWall):
        ite._area_columns()의 면적. Slab/Roof/Covering은 좌표기반 footprint 안분 적용.
        Door/Window/CurtainWall은 "시스템 전체 bounding치수(폭x높이, 두께제외)" 기반이라
        멀리언/패널 등 하위부품 단위가 아니라 문/창/커튼월 전체 대표 면적 하나로 계산되며
        (부품별로 안분할 대상이 아니라 안분 미적용, 전체값 그대로 사용).
      - 그 외 모든 클래스(Column/Beam/Member/Stair/Railing 등): 면적 없음(개수만) - 압출
        단면적이거나, 상세 메쉬일 경우 실제 면적이 아닌 표면적으로 잘못 계산될 위험이
        실측으로 확인되어 기본적으로 제외했다(잘못된 수치를 그럴듯하게 보여주는 것 방지).

    반환: {'by_class_split': {클래스: {구분라벨: {'count':int,'area':float|None}}},
           'by_class_total': {클래스: {'count':int,'area':float|None}}}
    """
    related = get_space_related_elements(ifc_file, space_entity)
    space_footprint = get_footprint_polygon(space_entity)

    split = defaultdict(lambda: defaultdict(lambda: {'count': 0, 'area': 0.0, '_has_area': False}))
    total = defaultdict(lambda: {'count': 0, 'area': 0.0, '_has_area': False})

    for e in related:
        cls = e.is_a()
        label, _reason = element_classification.get(e.GlobalId, ('판정불가', ''))

        area_val = None
        if cls in ('IfcWall', 'IfcWallStandardCase'):
            flat = ite._flatten_psets(e)
            v = flat.get('Qto_WallBaseQuantities.Gross_Side_Area')
            if isinstance(v, (int, float)):
                area_val, _method = _apportioned_area(ifc_file, e, space_entity, v, space_footprint)
        elif cls in _STRUCTURAL_AREA_MEANINGFUL_CLASSES:
            flat = ite._flatten_psets(e)
            cols = ite._area_columns(e, flat)
            if cols['면적(㎡)'] is not None:
                if cls in _AREA_APPORTION_CLASSES:
                    area_val, _method = _apportioned_area(ifc_file, e, space_entity, cols['면적(㎡)'], space_footprint)
                else:
                    area_val = cols['면적(㎡)']
        # 화이트리스트에 없는 클래스는 area_val=None 유지 -> 개수만 집계됨

        total[cls]['count'] += 1
        split[cls][label]['count'] += 1
        if area_val is not None:
            total[cls]['area'] += area_val
            total[cls]['_has_area'] = True
            split[cls][label]['area'] += area_val
            split[cls][label]['_has_area'] = True

    def _finalize(d):
        return {k: {'count': v['count'], 'area': round(v['area'], 2) if v['_has_area'] else None}
                for k, v in d.items()}

    return {
        'by_class_split': {cls: _finalize(labels) for cls, labels in split.items()},
        'by_class_total': _finalize(total),
    }


def _build_highlight_map(related, equipment, wall_classification):
    """평면도에서 색을 다르게 칠하기 위한 GlobalId -> 카테고리 매핑.
    카테고리: 'wall_internal'(내벽, 1차/2차 확정) / 'wall_internal_estimated'(내벽 추정-관계기반, 3차) /
              'wall_external_confirmed'(외벽, 1차/2차로 확정) /
              'wall_external_unknown'(판정불가로 외부 편입된 것 - 근거 없음, 구분 표시) /
              'related'(벽 이외 경계 관련 부재) / 'equipment'(조명/센서/소방설비)."""
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


# ===================================================================
# 5. Plotly 평면도 figure 생성
# ===================================================================

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

# 공간 클릭시 하이라이트 색상 (카테고리별로 뚜렷이 구분)
_HIGHLIGHT_COLORS = {
    'wall_internal': ('rgba(30,110,230,0.85)', 'rgba(15,70,160,1.0)'),              # 내벽(확정) = 진한 파랑
    'wall_internal_estimated': ('rgba(120,180,240,0.75)', 'rgba(60,120,190,1.0)'),  # 내벽(추정-관계기반) = 연한 파랑
    'wall_external_confirmed': ('rgba(230,90,30,0.85)', 'rgba(170,60,10,1.0)'),     # 외벽(판정됨) = 주황
    'wall_external_unknown': ('rgba(230,190,190,0.85)', 'rgba(160,50,50,1.0)'),     # 외벽(판정불가) = 연한 붉은/분홍(주황과 구분)
    'related': ('rgba(160,50,190,0.75)', 'rgba(110,20,140,1.0)'),                   # 벽 이외 관련부재 = 보라
}
_EQUIPMENT_COLOR = 'rgba(220,190,20,0.95)'  # 설비(조명/센서/소방) = 노랑 마커

_SPACE_FILL = 'rgba(100,180,120,0.35)'
_SPACE_FILL_SELECTED = 'rgba(230,100,60,0.55)'
_SPACE_LINE = 'rgba(60,140,80,0.9)'
_SPACE_LINE_SELECTED = 'rgba(200,60,20,1.0)'

_SPACE_UNMATCHED_FILL = 'rgba(190,190,190,0.35)'
_SPACE_UNMATCHED_LINE = 'rgba(150,150,150,0.8)'

# 공간 자동매핑 쌍 표시용 색상 팔레트 (양쪽 평면도에서 같은 번호는 항상 같은 색)
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
# 자동매핑은 됐지만 '지금 선택되지는 않은' 공간용 - 테두리 없이 배경만 아주 옅게(참고용
# 존재감만 표시). 선택된 공간은 벽 하이라이트(highlight_map)로 경계가 이미 뚜렷이
# 표시되므로, 매칭만 되고 선택 안 된 공간은 눈에 덜 띄게 해서 시각적으로 구분한다.
_PAIR_PALETTE_FAINT = [
    'rgba(31,119,180,0.12)', 'rgba(255,127,14,0.12)', 'rgba(44,160,44,0.12)',
    'rgba(214,39,40,0.12)', 'rgba(148,103,189,0.12)', 'rgba(140,86,75,0.12)',
    'rgba(227,119,194,0.12)', 'rgba(127,127,127,0.12)', 'rgba(188,189,34,0.12)',
    'rgba(23,190,207,0.12)',
]


def build_pair_labels(a_to_b):
    """match_spaces()가 반환한 a_to_b({A GlobalId: B GlobalId}) 딕셔너리로부터,
    양쪽 평면도에 표시할 번호 라벨을 한번에 생성한다 (같은 쌍은 항상 같은 번호).
    반환: (a_labels, b_labels) - 각각 {GlobalId: 번호} 딕셔너리."""
    a_labels, b_labels = {}, {}
    for i, (a_guid, b_guid) in enumerate(a_to_b.items(), start=1):
        a_labels[a_guid] = i
        if b_guid:
            b_labels[b_guid] = i
    return a_labels, b_labels


def build_plan_figure(plan_data, click_grid_spacing=0.5, selected_guid=None,
                       highlight_map=None, equipment_entities=None, pair_labels=None,
                       wall_segments=None):
    """plan_data(build_storey_plan_data 반환값)로 Plotly Figure 생성.

    Space는 내부에 보이지 않는 마커 격자를 깔아 '폴리곤 내부 아무 곳이나 클릭'해도
    선택되도록 한다(Plotly는 기본적으로 마커/점 단위로만 클릭을 인식하기 때문).

    selected_guid: 강조 표시할 선택된 Space의 GlobalId.
    highlight_map: build_space_detail()이 반환한 {GlobalId: 카테고리} 딕셔너리.
        선택된 Space와 관련된 구조요소를 카테고리별 색상으로 강조하고, 나머지 배경
        구조요소는 흐리게(faded) 처리해 '내부 객체(파랑)/외부 객체(주황)/기타 관련부재(보라)'가
        한눈에 구분되도록 한다. None이면(선택 없음) 기본 클래스별 색상으로 표시.
    wall_segments: build_space_detail()이 반환한 {벽 GlobalId: 클리핑된 Polygon|None}.
        벽을 강조할 때, 벽 전체가 아니라 '실제로 선택된 공간에 접한 부분만' 진하게
        칠하고, 벽의 나머지 부분(다른 공간에 접하거나 접하지 않는 부분)은 옅게(맥락용)
        표시한다 - 면적 계산(RelSpaceBoundary 기반 정밀 안분)과 화면 표시의 정밀도를
        맞추기 위함. 클리핑에 실패한 벽(None)은 기존처럼 벽 전체를 강조색으로 표시한다.
    equipment_entities: 선택된 Space '안에' 있는 설비(조명/센서/소방장치 등) ifcopenshell 엔티티
        목록. 평면도에 노란 마커로 추가 표시한다 (RelSpaceBoundary 대상이 아니라 별도로 그림).
    pair_labels: build_pair_labels()가 반환한 {GlobalId: 번호} 딕셔너리(이 평면도 쪽).
        주어지면(자동매핑 활성화시) 매칭된 공간마다 같은 번호 배지+같은 계열 색상을 칠해
        양쪽 평면도에서 어떤 공간끼리 매칭됐는지 클릭 없이도 한눈에 보이게 한다.
        매칭 안 된 공간은 회색으로 표시해 구분한다. None이면(자동매핑 비활성화) 기존
        기본 초록색 표시로 돌아간다.
    """
    import plotly.graph_objects as go
    fig = go.Figure()
    highlight_map = highlight_map or {}
    pair_labels = pair_labels or {}
    wall_segments = wall_segments or {}

    # 구조요소(벽/기둥/보/바닥 등)
    for el in plan_data['structural']:
        cat = highlight_map.get(el['guid'])
        is_wall = el['class'] in ('IfcWall', 'IfcWallStandardCase')
        segment_poly = wall_segments.get(el['guid']) if is_wall else None

        if highlight_map:
            if cat in _HIGHLIGHT_COLORS:
                if segment_poly is not None:
                    # 정밀 표시: 벽 전체는 맥락용으로 옅게, 실제 접한 구간만 진하게 덧그림
                    xs_full, ys_full = _polygon_xy_lists(el['polygon'])
                    fig.add_trace(go.Scatter(
                        x=xs_full, y=ys_full, mode='lines', fill='toself',
                        line=dict(width=0.5, color=_FADED_LINE), fillcolor=_FADED_COLOR,
                        hoverinfo='text', text=el.get('hover') or f"{el['class']} {el['name']}".strip(),
                        showlegend=False,
                    ))
                    fill_c, line_c = _HIGHLIGHT_COLORS[cat]
                    xs, ys = _polygon_xy_lists(segment_poly)
                    fig.add_trace(go.Scatter(
                        x=xs, y=ys, mode='lines', fill='toself',
                        line=dict(width=1.5, color=line_c), fillcolor=fill_c,
                        hoverinfo='text',
                        text=(el.get('hover') or f"{el['class']} {el['name']}".strip())
                             + '<br>(음영: 이 공간에 실제 접한 구간)',
                        showlegend=False,
                    ))
                    continue
                fill_c, line_c = _HIGHLIGHT_COLORS[cat]
                line_w = 1.5
            else:
                fill_c, line_c, line_w = _FADED_COLOR, _FADED_LINE, 0.5
        else:
            fill_c = _STRUCT_COLORS.get(el['class'], 'rgba(150,150,150,0.5)')
            line_c, line_w = 'rgba(60,60,60,0.6)', 0.5

        xs, ys = _polygon_xy_lists(el['polygon'])
        fig.add_trace(go.Scatter(
            x=xs, y=ys, mode='lines', fill='toself',
            line=dict(width=line_w, color=line_c), fillcolor=fill_c,
            hoverinfo='text',
            text=el.get('hover') or f"{el['class']} {el['name']}".strip(),
            showlegend=False,
        ))

    # Space: 시각적 채움(폴리곤) + 클릭 히트영역(격자 마커, 투명) + (선택) 매칭쌍 번호배지
    badge_x, badge_y, badge_text, badge_color, badge_line = [], [], [], [], []
    for sp in plan_data['spaces']:
        is_sel = (selected_guid is not None and sp['guid'] == selected_guid)
        pair_no = pair_labels.get(sp['guid'])

        if is_sel:
            # 선택된 공간: 자체 테두리를 그리지 않는다 - highlight_map으로 강조되는
            # 벽(내벽/외벽 등)의 윤곽이 이미 "실제 면적 산정에 쓰인 경계"를 보여주므로,
            # 공간 폴리곤 자체의 테두리는 그와 겹쳐 오히려 헷갈릴 수 있어 없앤다.
            fill_c, line_c, line_w = _SPACE_FILL_SELECTED, _SPACE_LINE_SELECTED, 0
        elif pair_no is not None:
            # 자동매핑은 됐지만 지금 선택되지는 않은 공간: 테두리 없이, 배경만 아주
            # 옅게(존재감만) 표시해 선택된 공간과 한눈에 구분되게 한다.
            idx = (pair_no - 1) % len(_PAIR_PALETTE)
            fill_c, line_c, line_w = _PAIR_PALETTE_FAINT[idx], _PAIR_PALETTE_LINE[idx], 0
        elif pair_labels:  # 자동매핑은 켜져있는데 이 공간은 매칭 안 됨
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

    # 설비(조명/센서/소방장치): 선택된 Space 안에 있는 것만 노란 마커로 표시
    if equipment_entities:
        ex, ey, etext = [], [], []
        for e in equipment_entities:
            poly = get_footprint_polygon(e)
            if poly is not None and not poly.is_empty:
                c = poly.centroid
                ex.append(c.x); ey.append(c.y)
            else:
                # 형상이 없으면 배치 좌표(ObjectPlacement)라도 사용
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

