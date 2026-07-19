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
from shapely.geometry import Polygon, Point, box, MultiPolygon
from shapely.ops import unary_union

import ifc_to_excel as ite  

_SETTINGS = geom.settings()
_SETTINGS.set('use-world-coords', True)

PLAN_STRUCTURAL_CLASSES = (
    'IfcWall', 'IfcWallStandardCase', 'IfcColumn', 'IfcBeam',
    'IfcSlab', 'IfcCurtainWall', 'IfcDoor', 'IfcWindow',
    'IfcRailing', 'IfcCovering', 'IfcStair', 'IfcStairFlight', 
    'IfcRamp', 'IfcRampFlight'
)


def load_ifc(path):
    ifc_file = ifcopenshell.open(path)

    storeys = []
    for s in ifc_file.by_type('IfcBuildingStorey'):
        storeys.append({'Name': s.Name, 'Elevation': s.Elevation, 'entity': s})
    storeys.sort(key=lambda x: (x['Elevation'] is None, x['Elevation']))

    wall_classification = ite._determine_wall_classification(ifc_file)
    element_classification = ite._determine_element_classification(ifc_file)

    return {
        'ifc_file': ifc_file,
        'storeys': storeys,
        'wall_classification': wall_classification,
        'element_classification': element_classification,
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
        verts = np.array(shape.geometry.verts).reshape(-1, 3)
        faces = np.array(shape.geometry.faces).reshape(-1, 3)
        if len(verts) > 0 and len(faces) > 0:
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
            if polys:
                u = unary_union(polys)
                if not u.is_empty:
                    return u
    except Exception:
        pass

    child_polys = []
    for rel in getattr(ent, 'IsDecomposedBy', []):
        for child in rel.RelatedObjects:
            child_poly = get_footprint_polygon(child, tol)
            if child_poly is not None and not child_poly.is_empty:
                child_polys.append(child_poly)
                
    if child_polys:
        try:
            u = unary_union(child_polys)
            if not u.is_empty:
                return u
        except Exception:
            pass
            
    return None


_element_footprint_cache = {}


def get_footprint_polygon_cached(ent):
    """[면적 산정 중복계산 방지] 벽/슬래브 등 부재의 footprint는 공간과 무관하게
    부재 자체에 고정된 값인데, 한 부재가 여러 공간에 걸쳐 있으면(예: 내벽이 양쪽
    공간에서 각각 조회됨) 안분·판정 과정에서 같은 부재의 footprint가 공간 수만큼
    반복 계산되고 있었다. GlobalId 기준으로 한 번만 계산해 재사용한다."""
    key = ent.GlobalId
    if key not in _element_footprint_cache:
        _element_footprint_cache[key] = get_footprint_polygon(ent)
    return _element_footprint_cache[key]


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


def get_elements_for_storey(storey_entity, classes=None, max_decompose_depth=3):
    storey_ifc = storey_entity['entity'] if isinstance(storey_entity, dict) else storey_entity
    elements = []
    seen_guids = set()

    def _add(el):
        if el.GlobalId in seen_guids:
            return
        seen_guids.add(el.GlobalId)
        if classes is None or el.is_a() in classes:
            elements.append(el)

    def _collect_children(el, depth, is_curtain_wall=False):
        if depth >= max_decompose_depth:
            return
        for rel in (getattr(el, 'IsDecomposedBy', None) or []):
            for child in rel.RelatedObjects:
                # 커튼월 내부에 포함된 하위 부품은 독립 개체로 추가하지 않음 (중복 방지)
                if not is_curtain_wall:
                    _add(child)
                _collect_children(child, depth + 1, is_curtain_wall or el.is_a('IfcCurtainWall'))

    top_level = []
    for rel in (storey_ifc.ContainsElements or []):
        top_level.extend(rel.RelatedElements)
    for rel in (storey_ifc.IsDecomposedBy or []):
        top_level.extend(rel.RelatedObjects)

    for el in top_level:
        _add(el)
        _collect_children(el, 0, el.is_a('IfcCurtainWall'))

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


def build_storey_plan_data(storey_entity, tol=0.05, wall_classification=None, ifc_file=None):
    spaces_raw = get_elements_for_storey(storey_entity, classes={'IfcSpace'})
    structural_raw = get_elements_for_storey(storey_entity, classes=set(PLAN_STRUCTURAL_CLASSES))

    cached_footprints = {}
    if ifc_file is not None:
        storey_ifc = storey_entity['entity'] if isinstance(storey_entity, dict) else storey_entity
        cache_key = (id(ifc_file), storey_ifc.GlobalId)
        for el, fp in _storey_candidate_footprint_cache.get(cache_key, []):
            cached_footprints[el.GlobalId] = fp

    spaces = []
    for sp in spaces_raw:
        poly = get_footprint_polygon(sp, tol=tol)
        if poly is None:
            continue
        spaces.append({'guid': sp.GlobalId, 'name': sp.Name or '(이름없음)', 'polygon': poly, 'entity': sp})

    structural = []
    for el in structural_raw:
        poly = cached_footprints.get(el.GlobalId)
        if poly is None:
            poly = get_footprint_polygon(el, tol=tol)
        if poly is None:
            continue
        hover = get_element_hover_info(el, wall_classification=wall_classification)
        structural.append({'guid': el.GlobalId, 'class': el.is_a(), 'name': el.Name or '',
                            'polygon': poly, 'hover': hover, 'entity': el})

    return {'spaces': spaces, 'structural': structural}


EQUIPMENT_CLASSES = ('IfcLightFixture', 'IfcSensor', 'IfcFireSuppressionTerminal', 'IfcAlarm')


def _find_space_storey(space_entity):
    for rel in (space_entity.Decomposes or []):
        obj = getattr(rel, 'RelatingObject', None)
        if obj is not None and obj.is_a('IfcBuildingStorey'):
            return obj
    return None


_storey_candidate_footprint_cache = {}  


def _get_storey_candidate_footprints(ifc_file, storey):
    key = (id(ifc_file), storey.GlobalId)
    if key not in _storey_candidate_footprint_cache:
        candidates = get_elements_for_storey(storey, classes=set(ite.ELEMENT_CLASSIFICATION_TARGET_CLASSES))
        pairs = []
        for el in candidates:
            fp = get_footprint_polygon(el)
            if fp is not None and not fp.is_empty:
                pairs.append((el, fp))
        _storey_candidate_footprint_cache[key] = pairs
    return _storey_candidate_footprint_cache[key]


_storey_wall_candidate_footprint_cache = {}


def _get_storey_wall_candidate_footprints(ifc_file, storey):
    """[수정사항] ELEMENT_CLASSIFICATION_TARGET_CLASSES에는 IfcWall이 빠져있어, 벽은
    RelSpaceBoundary 관계가 있을 때만 공간에 연결되고 관계가 0건인 벽(예: 조립체 병합으로
    승격은 됐지만 실제 관계는 없는 코어층)은 어떤 공간의 related 목록에도 절대 못
    들어가는 문제가 실측으로 확인됨(공간별 합산에서 통째로 누락). 벽 전용 후보 목록을
    별도로 캐싱해 지오메트리 폴백 대상에 포함한다."""
    key = (id(ifc_file), storey.GlobalId)
    if key not in _storey_wall_candidate_footprint_cache:
        candidates = get_elements_for_storey(storey, classes={'IfcWall', 'IfcWallStandardCase'})
        pairs = []
        for el in candidates:
            fp = get_footprint_polygon_cached(el)
            if fp is not None and not fp.is_empty:
                pairs.append((el, fp))
        _storey_wall_candidate_footprint_cache[key] = pairs
    return _storey_wall_candidate_footprint_cache[key]


def precompute_storey_geometry(ifc_file, storeys, status_cb=None):
    total = len(storeys)
    for i, storey in enumerate(storeys, start=1):
        if status_cb:
            status_cb(storey['Name'], i, total)
        _get_storey_candidate_footprints(ifc_file, storey['entity'])


def get_space_related_elements(ifc_file, space_entity, geometric_fallback=True, adjacency_tol=0.15,
                                wall_adjacency_tol=0.35):
    """[수정사항] 벽 전용 지오메트리 폴백의 허용거리를 일반 부재(0.15m)와 분리해 0.35m로
    완화한다. 실측 결과, 관계가 0건인 마감층 벽들이 공간 footprint로부터 0.20~0.30m
    떨어져 있어 기존 0.15m 허용거리로는 못 잡히는 사례가 다수 확인됨(층별로 최대
    983㎡, 벽 27개가 이 사유로 누락). 0.35m는 실측된 정상 레이어 간격(0.05~0.3m) 범위를
    안전하게 포괄하면서, 이 값은 개별 벽 하나의 footprint 단위로 적용되므로(등방 버퍼를
    긴 벽 전체나 조립체 전체에 씌우는 것과 달리) 이전에 겪었던 '옆방 오검출' 부작용
    위험은 낮다."""
    by_guid = {}
    for rel in ifc_file.by_type('IfcRelSpaceBoundary'):
        if rel.RelatingSpace == space_entity and rel.RelatedBuildingElement is not None:
            el = rel.RelatedBuildingElement
            by_guid[el.GlobalId] = el

    if geometric_fallback:
        storey = _find_space_storey(space_entity)
        if storey is not None:
            space_fp = get_footprint_polygon(space_entity)
            if space_fp is not None and not space_fp.is_empty:
                for el, el_fp in _get_storey_candidate_footprints(ifc_file, storey):
                    if el.GlobalId in by_guid:
                        continue
                    if space_fp.distance(el_fp) <= adjacency_tol:
                        by_guid[el.GlobalId] = el
                # 벽 전용 폴백 (기존에는 벽이 이 지오메트리 폴백 대상에서 전부 빠져있었음)
                for el, el_fp in _get_storey_wall_candidate_footprints(ifc_file, storey):
                    if el.GlobalId in by_guid:
                        continue
                    if space_fp.distance(el_fp) <= wall_adjacency_tol:
                        by_guid[el.GlobalId] = el

    return list(by_guid.values())


_equipment_index_cache = {}

def _get_equipment_index(ifc_file):
    key = id(ifc_file)
    if key not in _equipment_index_cache:
        index = defaultdict(list)
        for rel in ifc_file.by_type('IfcRelContainedInSpatialStructure'):
            sp = rel.RelatingStructure
            if sp is not None and sp.is_a('IfcSpace'):
                for el in rel.RelatedElements:
                    index[sp.GlobalId].append(el)
        _equipment_index_cache[key] = index
    return _equipment_index_cache[key]


def get_space_contained_equipment(ifc_file, space_entity, classes=EQUIPMENT_CLASSES):
    index = _get_equipment_index(ifc_file)
    equipment = []
    for el in index.get(space_entity.GlobalId, []):
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

_STRUCTURAL_AREA_MEANINGFUL_CLASSES = {'IfcSlab', 'IfcRoof', 'IfcCovering', 'IfcDoor', 'IfcWindow', 'IfcCurtainWall'}

_boundary_index_cache = {}  

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
    a, b = np.array(seg_p1), np.array(seg_p2)
    direction = b - a
    length = np.linalg.norm(direction)
    if length < 1e-6:
        return None
    x_axis = direction / length
    y_axis = np.array([-x_axis[1], x_axis[0]])  
    rel = np.array(local_xy) - a
    return '+' if np.dot(rel, y_axis) >= 0 else '-'


def _space_side_of_member(space_entity, member_m_inv, axis_segments=None):
    poly = get_footprint_polygon(space_entity)
    if poly is None or poly.is_empty:
        return None
    cx, cy = poly.centroid.x * 1000.0, poly.centroid.y * 1000.0
    try:
        import ifcopenshell.util.placement as plc
        space_m = plc.get_local_placement(space_entity.ObjectPlacement)
    except Exception:
        return None
    cz = space_m[2, 3]  
    local_pt = member_m_inv @ np.array([cx, cy, cz, 1.0])

    if axis_segments:
        idx = _nearest_segment_index((local_pt[0], local_pt[1]), axis_segments)
        p1, p2 = axis_segments[idx]
        side = _segment_side((local_pt[0], local_pt[1]), p1, p2)
        if side is not None:
            return side

    return '+' if local_pt[1] >= 0 else '-'


def _relspaceboundary_precise_areas(ifc_file, member_entity, own_side_area_m2, tol_factor=1.10):
    import ifcopenshell.util.placement as plc

    try:
        member_m = plc.get_local_placement(member_entity.ObjectPlacement)
        member_m_inv = np.linalg.inv(member_m)
    except Exception:
        return None

    axis_segments = _member_axis_segments(member_entity)  

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
        return None  

    per_space = defaultdict(float)
    for b in boundaries:
        per_space[b['space_guid']] += b['area_m2']
    return dict(per_space)


def _apportioned_area(ifc_file, member_entity, target_space_entity, own_side_area_m2, space_footprint,
                       tol_factor=1.10, segment_polygon=None, wall_footprint_polygon=None):
    precise = _relspaceboundary_precise_areas(ifc_file, member_entity, own_side_area_m2, tol_factor)
    if precise is not None and target_space_entity.GlobalId in precise:
        return precise[target_space_entity.GlobalId], '정밀(RelSpaceBoundary)'

    if member_entity.is_a('IfcWall'):
        wfp = wall_footprint_polygon if wall_footprint_polygon is not None else get_footprint_polygon_cached(member_entity)
        if segment_polygon is None:
            seg_result = get_space_wall_segment_polygon(ifc_file, member_entity, target_space_entity, wfp)
            segment_polygon = seg_result[0] if seg_result is not None else None
        if segment_polygon is not None and wfp is not None and wfp.area > 0:
            fraction = segment_polygon.area / wfp.area
            return own_side_area_m2 * fraction, '정밀(화면표시와 동일 클리핑)'

    fraction = _space_portion_fraction(member_entity, space_footprint)
    if fraction is None:
        return own_side_area_m2, '실패-전체값사용(과다산정 가능)'
    return own_side_area_m2 * fraction, '근사(footprint버퍼)'


def _space_portion_fraction(member_entity, space_footprint, buffer_dist=_APPORTION_BUFFER_M):
    if space_footprint is None or space_footprint.is_empty:
        return None
    member_footprint = get_footprint_polygon_cached(member_entity)
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


def _polygon_edges(poly):
    if poly.geom_type == 'MultiPolygon':
        polys = list(poly.geoms)
    elif poly.geom_type == 'Polygon':
        polys = [poly]
    else:
        return [] 
    edges = []
    for p in polys:
        coords = list(p.exterior.coords)
        edges.extend((coords[i], coords[i + 1]) for i in range(len(coords) - 1))
    return edges


def _collinear_overlap_segment(seg1, seg2, line_tol=0.02, min_overlap=0.05):
    (x1, y1), (x2, y2) = seg1
    (x3, y3), (x4, y4) = seg2
    d1 = np.array([x2 - x1, y2 - y1])
    len1 = np.linalg.norm(d1)
    if len1 < 1e-6:
        return None
    d1n = d1 / len1

    d3 = np.array([x4 - x3, y4 - y3])
    len3 = np.linalg.norm(d3)
    if len3 < 1e-6:
        return None
    d3n = d3 / len3

    if abs(d1n[0] * d3n[1] - d1n[1] * d3n[0]) > 0.02: 
        return None

    v = np.array([x3 - x1, y3 - y1])
    if abs(v[0] * d1n[1] - v[1] * d1n[0]) > line_tol:  
        return None

    t3_start = np.dot(np.array([x3 - x1, y3 - y1]), d1n)
    t3_end = np.dot(np.array([x4 - x1, y4 - y1]), d1n)
    t_min = max(0.0, min(t3_start, t3_end))
    t_max = min(len1, max(t3_start, t3_end))
    if t_max - t_min < min_overlap:
        return None

    p_start = np.array([x1, y1]) + d1n * t_min
    p_end = np.array([x1, y1]) + d1n * t_max
    return (tuple(p_start), tuple(p_end))


def _space_wall_edge_overlap_local_points(wall_footprint, space_footprint, member_m_inv):
    wall_edges = _polygon_edges(wall_footprint)
    space_edges = _polygon_edges(space_footprint)

    local_pts = []
    for we in wall_edges:
        for se in space_edges:
            overlap = _collinear_overlap_segment(we, se)
            if overlap is None:
                continue
            for (wx, wy) in overlap:
                local_pt = member_m_inv @ np.array([wx * 1000.0, wy * 1000.0, 0.0, 1.0])
                local_pts.append((local_pt[0], local_pt[1]))

    return local_pts if local_pts else None


def get_local_thickness_window(wall_entity, seg_poly, thickness_reach=0.4, length_margin=0.0):
    """세그먼트 폴리곤(공간 하나와 접하는 벽의 국소 조각)을 감싸는 윈도우를 만들되,
    벽의 로컬 좌표계 기준으로 '길이 방향(로컬 X)'은 세그먼트 범위 + 작은 여유만 두고,
    '두께/법선 방향(로컬 Y)'만 넉넉히(기본 0.6m) 확장한다.
    [이유] 단순히 polygon.buffer(1.0)처럼 등방(모든 방향 동일)으로 키우면, 복도를 따라
    이어지는 옆방(같은 쪽에 나란히 있는 다음 공간)까지 창이 넓어져 버그가 생긴다
    (실측 확인: P-04 옆 P-03이 버퍼에 걸려 반대편 공간으로 오인됨). 벽의 실제 두께
    방향으로만 확장해야 '반대편'을 정확히 검사할 수 있다."""
    import ifcopenshell.util.placement as plc
    try:
        m = plc.get_local_placement(wall_entity.ObjectPlacement)
        m_inv = np.linalg.inv(m)
    except Exception:
        return seg_poly.buffer(thickness_reach)

    def to_local(pt):
        world = np.array([pt[0] * 1000.0, pt[1] * 1000.0, 0.0, 1.0])
        local = m_inv @ world
        return local[0], local[1]

    geoms = seg_poly.geoms if seg_poly.geom_type == 'MultiPolygon' else [seg_poly]
    local_pts = []
    for g in geoms:
        local_pts.extend(to_local(p) for p in g.exterior.coords)
    if not local_pts:
        return seg_poly.buffer(thickness_reach)

    xs = [p[0] for p in local_pts]
    ys = [p[1] for p in local_pts]
    length_margin_mm = length_margin * 1000.0
    thickness_reach_mm = thickness_reach * 1000.0
    # 관례상 로컬 X=벽 길이축, 로컬 Y=벽 두께(법선)축(이 코드베이스 전반의 기존 관례와 동일)
    x_min, x_max = min(xs) - length_margin_mm, max(xs) + length_margin_mm
    y_min, y_max = min(ys) - thickness_reach_mm, max(ys) + thickness_reach_mm

    local_corners = [(x_min, y_min), (x_max, y_min), (x_max, y_max), (x_min, y_max)]
    world_pts = []
    for (lx, ly) in local_corners:
        w = m @ np.array([lx, ly, 0.0, 1.0])
        world_pts.append((w[0] / 1000.0, w[1] / 1000.0))
    try:
        win = Polygon(world_pts)
        return win if win.is_valid else seg_poly.buffer(thickness_reach)
    except Exception:
        return seg_poly.buffer(thickness_reach)


def get_space_wall_segment_polygon(ifc_file, wall_entity, space_entity, wall_footprint_polygon, space_footprint=None):
    import ifcopenshell.util.placement as plc

    try:
        member_m = plc.get_local_placement(wall_entity.ObjectPlacement)
        member_m_inv = np.linalg.inv(member_m)
    except Exception:
        return None

    if space_footprint is None:
        space_footprint = get_footprint_polygon_cached(space_entity)

    rels = _get_boundary_index(ifc_file).get(wall_entity.GlobalId, [])
    local_pts = []
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
            local_pts.append((local_pt[0], local_pt[1]))

    method = 'ConnectionGeometry' if local_pts else None

    if not local_pts:
        if space_footprint is not None and wall_footprint_polygon is not None:
            edge_pts = _space_wall_edge_overlap_local_points(
                wall_footprint_polygon, space_footprint, member_m_inv)
            if edge_pts is not None:
                local_pts = edge_pts
                method = 'edge겹침추정'

    if not local_pts:
        return None  

    margin = 50.0  
    xs = [p[0] for p in local_pts]
    x_min, x_max = min(xs) - margin, max(xs) + margin

    if wall_footprint_polygon is None or wall_footprint_polygon.is_empty:
        return None

    def _ring_to_local(coords):
        return [tuple(member_m_inv @ np.array([x * 1000.0, y * 1000.0, 0.0, 1.0]))[:2] for (x, y) in coords]

    if wall_footprint_polygon.geom_type == 'Polygon':
        frag_polys = [wall_footprint_polygon]
    elif wall_footprint_polygon.geom_type == 'MultiPolygon':
        frag_polys = list(wall_footprint_polygon.geoms)
    else:
        return None

    local_frags = []
    for fp in frag_polys:
        try:
            lp = Polygon(_ring_to_local(fp.exterior.coords))
        except Exception:
            continue
        if lp.is_valid and not lp.is_empty:
            local_frags.append(lp)
    if not local_frags:
        return None
    local_poly = local_frags[0] if len(local_frags) == 1 else MultiPolygon(local_frags)

    minx, miny, maxx, maxy = local_poly.bounds
    clip_box = box(x_min, miny - margin, x_max, maxy + margin)
    clipped_local = local_poly.intersection(clip_box)
    if clipped_local.is_empty:
        return None

    def _to_world_xy(pt):
        world_pt = member_m @ np.array([pt[0], pt[1], 0.0, 1.0])
        return (world_pt[0] / 1000.0, world_pt[1] / 1000.0)

    if clipped_local.geom_type == 'Polygon':
        clipped_frags = [clipped_local]
    elif clipped_local.geom_type == 'MultiPolygon':
        clipped_frags = list(clipped_local.geoms)
    elif clipped_local.geom_type == 'GeometryCollection':
        clipped_frags = [g for g in clipped_local.geoms if g.geom_type == 'Polygon']
    else:
        return None
    if not clipped_frags:
        return None

    world_polys = []
    for frag in clipped_frags:
        world_coords = [_to_world_xy(p) for p in frag.exterior.coords]
        wp = Polygon(world_coords)
        if wp.is_valid and not wp.is_empty:
            world_polys.append(wp)
    if not world_polys:
        return None

    proximity_buffer_m = 0.5  
    if space_footprint is not None:
        trim_region = space_space_footprint = space_footprint.buffer(proximity_buffer_m) if space_footprint else None
        trimmed = []
        for wp in world_polys:
            t = wp.intersection(trim_region)
            if t.is_empty:
                continue
            if t.geom_type == 'Polygon':
                trimmed.append(t)
            elif t.geom_type == 'MultiPolygon':
                trimmed.extend(g for g in t.geoms if g.geom_type == 'Polygon' and not g.is_empty)
        if not trimmed:
            return None  
        world_polys = trimmed

    result = world_polys[0] if len(world_polys) == 1 else MultiPolygon(world_polys)
    return result, method


_wall_side_area_cache = {}


def _get_wall_side_area_m2(ent, flat_props=None):
    """[수정사항] Qto_WallBaseQuantities.Gross_Side_Area는 신뢰도 문제로 더 이상
    사용하지 않고 좌표 기반 bounding치수 계산만 사용한다. 또한 벽 하나가 여러
    공간에 걸쳐 있으면(내벽) 공간마다 반복 호출되므로, 벽 자체의 '전체 옆면적'은
    GlobalId 기준으로 한 번만 계산해 캐시한다(공간별로 달라지는 것은 이후 안분
    비율뿐이며, 안분 비율 계산은 이 캐시와 무관하게 공간별로 그대로 수행된다)."""
    key = ent.GlobalId
    if key in _wall_side_area_cache:
        return _wall_side_area_cache[key]
    raw_area, method = ite._get_system_bounding_face_area(ent)
    if raw_area is not None:
        result = (raw_area * 1e-6, f'폴백-{method}')
    else:
        result = (None, None)
    _wall_side_area_cache[key] = result
    return result


_wall_space_split_cache = {}


def compute_wall_space_splits(ifc_file, storey_entity):
    """[신규] 층 안의 모든 벽을, 접한 공간들에게 '배분액의 합 = 벽 자신의 총면적'이 되도록
    미리 분할해둔다(층 단위로 정확히 한 번 계산). 공간별 조회(build_space_detail 등)는
    이 결과에서 (벽,공간) 키로 값을 찾아오기만 하면 되므로, 여러 공간에 걸친 벽을 각
    공간에 전액씩 중복 부여하던 문제가 애초에 발생하지 않는다(총량 = 분할합산 + 잔여
    등식이 근사가 아니라 항상 정확히 성립).
    분배 비율은 기존 _apportioned_area()가 이미 계산하는 세그먼트/footprint 기반
    기하학적 가중치를 그대로 재사용하되, 공간들에 대한 값의 합이 own_area가 되도록
    정규화(rescale)한다 - 그래서 단순 균등분할보다, 각 공간과 실제로 맞닿은 구간의
    비중(예: 여러 방을 지나는 긴 복도벽에서 방마다 접한 길이가 다른 경우)을 더 정확히
    반영한다. 기하 계산이 전부 실패하면 균등분할로 폴백한다.
    반환: {(wall_guid, space_guid): 배분된_면적(㎡)}"""
    import sys
    _self = sys.modules[__name__]
    walls_by_storey = ite._wall_footprints_by_storey(ifc_file, _self)
    walls_here = walls_by_storey.get(storey_entity.GlobalId, [])
    wall_fp_by_guid = {w.GlobalId: fp for w, fp in walls_here}

    spaces = get_elements_for_storey(storey_entity, classes={'IfcSpace'})
    space_fp_by_guid = {sp.GlobalId: get_footprint_polygon_cached(sp) for sp in spaces}

    wall_to_spaces = defaultdict(list)
    for sp in spaces:
        related = get_space_related_elements(ifc_file, sp)
        for e in related:
            if e.is_a('IfcWall'):
                wall_to_spaces[e.GlobalId].append(sp)

    splits = {}
    for w, _fp in walls_here:
        touching = wall_to_spaces.get(w.GlobalId, [])
        if not touching:
            continue
        own, _src = _get_wall_side_area_m2(w)
        if own is None or own <= 0:
            continue
        w_fp = wall_fp_by_guid.get(w.GlobalId)

        if len(touching) == 1:
            splits[(w.GlobalId, touching[0].GlobalId)] = round(own, 4)
            continue

        raw = {}
        for sp in touching:
            sp_fp = space_fp_by_guid.get(sp.GlobalId)
            val, _method = _apportioned_area(ifc_file, w, sp, own, sp_fp, wall_footprint_polygon=w_fp)
            raw[sp.GlobalId] = max(val, 0.0)
        total_raw = sum(raw.values())
        if total_raw <= 0:
            n = len(touching)
            for sp in touching:
                splits[(w.GlobalId, sp.GlobalId)] = round(own / n, 4)
        else:
            for sp in touching:
                splits[(w.GlobalId, sp.GlobalId)] = round(own * (raw[sp.GlobalId] / total_raw), 4)

    return splits


def _get_wall_space_splits_cached(ifc_file, storey_entity):
    key = (id(ifc_file), storey_entity.GlobalId)
    if key not in _wall_space_split_cache:
        _wall_space_split_cache[key] = compute_wall_space_splits(ifc_file, storey_entity)
    return _wall_space_split_cache[key]


def build_space_detail(ifc_file, wall_classification, space_entity):
    related = get_space_related_elements(ifc_file, space_entity)
    equipment = get_space_contained_equipment(ifc_file, space_entity)
    space_footprint = get_footprint_polygon_cached(space_entity)

    wall_segment_polygons = {}
    wall_segment_methods = {}
    wall_footprints = {}
    for e in related:
        if not e.is_a('IfcWall'):
            continue
        wall_footprint = get_footprint_polygon_cached(e)
        wall_footprints[e.GlobalId] = wall_footprint
        result = get_space_wall_segment_polygon(ifc_file, e, space_entity, wall_footprint,
                                                 space_footprint=space_footprint)
        if result is not None:
            wall_segment_polygons[e.GlobalId], wall_segment_methods[e.GlobalId] = result
        else:
            wall_segment_polygons[e.GlobalId] = None

    wall_segment_stats = {
        'precise_cg': sum(1 for m in wall_segment_methods.values() if m == 'ConnectionGeometry'),
        'precise_edge': sum(1 for m in wall_segment_methods.values() if m == 'edge겹침추정'),
        'fallback': sum(1 for v in wall_segment_polygons.values() if v is None),
    }

    class_counts = Counter(e.is_a() for e in related)

    # [수정사항] 벽마다 그때그때 _apportioned_area()로 안분하면(내벽의 경우 접한 각
    # 공간에 전액에 가까운 값을 부여) 여러 공간에 걸친 벽의 면적이 중복 집계되어 "층
    # 전체 총량 = 공간별 분할합산 + 잔여"라는 등식이 깨졌었다. 층 단위로 미리 정확히
    # 분할해둔 표(compute_wall_space_splits, 합=own_area 보장)를 조회만 하도록 바꿔
    # 이 등식이 항상 성립하게 한다.
    storey_entity = _find_space_storey(space_entity)
    wall_splits = _get_wall_space_splits_cached(ifc_file, storey_entity) if storey_entity is not None else {}

    wall_simple_counts = Counter()
    wall_simple_area = Counter()
    wall_detail_counts = Counter()
    wall_area_apportioned = True  
    for e in related:
        if not e.is_a('IfcWall'):
            continue
        # [수정사항] 벽 전체(GlobalId) 라벨만 쓰면, '혼합(내/외벽 복합)'으로 판정된 벽이
        # 접하는 모든 공간에서 일괄적으로 '외부(판정됨)'으로 집계되어(실제로는 일부
        # 공간에서는 내벽인데도) 여러 공간에 걸쳐 면적이 중복 과다산정되는 문제가 실측으로
        # 확인됨(예: 9개 공간과 접한 566㎡ 벽 하나가 9곳 전부에서 외벽으로 잡힘). 이 공간
        # 전용 pair-level 라벨이 있으면 그것을 우선 사용하고, 없을 때만(관계 자체가 없던
        # 벽 등) 벽 전체 라벨로 폴백한다.
        result, _reason = wall_classification.get(
            (e.GlobalId, space_entity.GlobalId), wall_classification.get(e.GlobalId, ('판정불가', '')))
        simple, detail_label = _wall_display_category(result)
        wall_simple_counts[simple] += 1
        wall_detail_counts[detail_label] += 1
        area_val = wall_splits.get((e.GlobalId, space_entity.GlobalId))
        if area_val is not None:
            wall_simple_area[simple] += area_val
        else:
            wall_area_apportioned = False

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

    equipment_counts = Counter(e.is_a() for e in equipment)

    # [수정사항] 공간 면적도 Qto(NetFloorArea/GrossFloorArea) 우선 사용을 중단하고,
    # 다른 부재와 동일하게 좌표 기반 footprint 계산(_area_columns)으로 통일한다.
    # (flat_sp는 좌표계산이 실패할 경우의 최후 폴백에서만 참고됨)
    flat_sp = ite._flatten_psets(space_entity)
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
        'highlight_map': _build_highlight_map(related, equipment, wall_classification,
                                               space_guid=space_entity.GlobalId),
        'wall_segment_polygons': wall_segment_polygons,
        'wall_segment_stats': wall_segment_stats,
    }


def compute_wall_area_attribution(ifc_file, storey_entity):
    """벽체의 층단위 총면적을 '공간과 접해 분할된 부분(matched)'과 '공간과 무관해
    잘리지 않은 잔여 부분(residual)'으로 나눈다.
    [설계 의도] matched를 residual로부터 total - residual로 정의하므로
    total = matched + residual이 근사가 아니라 항상 정확히 성립한다 - 이 값이 안
    맞으면 그건 계산 버그이지 '원래 그런 데이터라서 어쩔 수 없는 오차'가 아님을
    명확히 구분하기 위함. residual은 어느 공간의 related 목록에도 전혀 안 걸리는
    벽(관계 0건 + 지오메트리 폴백도 실패)의 own_area 합계이며, 이런 벽은 대개
    방으로 모델링되지 않은 영역(주차장·설비공간 등)에 있거나, 드물게는 폴백
    허용거리(0.35m)를 벗어난 마감층일 수 있다.
    반환: {'total','matched','residual','residual_count','residual_walls':[(name,guid,area),...]}"""
    import sys
    _self = sys.modules[__name__]
    walls_by_storey = ite._wall_footprints_by_storey(ifc_file, _self)
    walls_here = [w for w, _fp in walls_by_storey.get(storey_entity.GlobalId, [])]

    spaces = get_elements_for_storey(storey_entity, classes={'IfcSpace'})
    covered = set()
    for sp in spaces:
        related = get_space_related_elements(ifc_file, sp)
        covered.update(e.GlobalId for e in related if e.is_a('IfcWall'))

    total = 0.0
    residual = 0.0
    residual_walls = []
    for w in walls_here:
        v, _src = _get_wall_side_area_m2(w)
        if v is None:
            continue
        total += v
        if w.GlobalId not in covered:
            residual += v
            residual_walls.append((w.Name or '(이름없음)', w.GlobalId, round(v, 2)))

    residual_walls.sort(key=lambda x: -x[2])
    return {
        'total': round(total, 2),
        'matched': round(total - residual, 2),
        'residual': round(residual, 2),
        'residual_count': len(residual_walls),
        'residual_walls': residual_walls,
    }


def build_space_structural_breakdown(ifc_file, element_classification, wall_classification, space_entity):
    related = get_space_related_elements(ifc_file, space_entity)
    space_footprint = get_footprint_polygon_cached(space_entity)

    # 벽은 build_space_detail과 동일하게, 층 단위로 미리 분할해둔 표를 조회한다
    # (합=own_area 보장 -> 총량=분할합산+잔여 등식이 항상 성립).
    storey_entity = _find_space_storey(space_entity)
    wall_splits = _get_wall_space_splits_cached(ifc_file, storey_entity) if storey_entity is not None else {}

    split = defaultdict(lambda: defaultdict(lambda: {'count': 0, 'area': 0.0, '_has_area': False}))
    total = defaultdict(lambda: {'count': 0, 'area': 0.0, '_has_area': False})

    for e in related:
        cls = e.is_a()
        # [수정사항 1] 벽은 element_classification(비-벽 부재 전용 분류)에 아예 없어 항상
        # '판정불가'로 잡히던 별도 버그가 있었다 - 벽 전용 wall_classification을 사용한다.
        # [수정사항 2] 벽 전체(GlobalId) 라벨만 쓰면 '혼합' 벽이 접하는 모든 공간에서
        # 일괄 외부로 집계되는 문제가 있었다 - (벽,공간) pair-level 라벨을 우선 사용한다.
        if cls in ('IfcWall', 'IfcWallStandardCase'):
            label, _reason = wall_classification.get(
                (e.GlobalId, space_entity.GlobalId), wall_classification.get(e.GlobalId, ('판정불가', '')))
        else:
            label, _reason = element_classification.get(e.GlobalId, ('판정불가', ''))

        area_val = None
        if cls in ('IfcWall', 'IfcWallStandardCase'):
            area_val = wall_splits.get((e.GlobalId, space_entity.GlobalId))
        elif cls in _STRUCTURAL_AREA_MEANINGFUL_CLASSES:
            flat = ite._flatten_psets(e)
            cols = ite._area_columns(e, flat)
            if cols['면적(㎡)'] is not None:
                if cls in _AREA_APPORTION_CLASSES:
                    area_val, _method = _apportioned_area(ifc_file, e, space_entity, cols['면적(㎡)'], space_footprint)
                else:
                    area_val = cols['면적(㎡)']

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


def _build_highlight_map(related, equipment, wall_classification, space_guid=None):
    hl = {}
    for e in related:
        if e.is_a('IfcWall'):
            # space_guid가 주어지면(특정 공간 조회 컨텍스트) pair-level 라벨을 우선 사용해
            # wall_simple_area 집계와 하이라이트 색상이 어긋나지 않도록 한다.
            if space_guid is not None:
                result, _ = wall_classification.get((e.GlobalId, space_guid),
                                                     wall_classification.get(e.GlobalId, ('판정불가', '')))
            else:
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
            hl[e.GlobalId] = e.is_a()
    for e in equipment:
        hl[e.GlobalId] = e.is_a()
    return hl


# ===================================================================
# 5. Plotly 평면도 figure 생성
# ===================================================================

_STRUCT_COLORS = {
    'IfcWall': 'rgba(90,90,90,0.35)',
    'IfcWallStandardCase': 'rgba(90,90,90,0.35)',
    'IfcColumn': 'rgba(40,40,40,0.35)',
    'IfcBeam': 'rgba(120,90,60,0.3)',
    'IfcSlab': 'rgba(200,190,170,0.2)',
    'IfcCurtainWall': 'rgba(120,170,220,0.25)',
    'IfcDoor': 'rgba(150,100,50,0.25)',
    'IfcWindow': 'rgba(120,200,230,0.25)',
    'IfcRailing': 'rgba(170,170,60,0.25)',
    'IfcCovering': 'rgba(90,160,150,0.25)',
    'IfcStair': 'rgba(200,90,150,0.25)',
    'IfcStairFlight': 'rgba(200,90,150,0.2)',
    'IfcRamp': 'rgba(200,90,150,0.2)',
    'IfcRampFlight': 'rgba(200,90,150,0.15)',
}
_FADED_COLOR = 'rgba(210,210,210,0.18)'
_FADED_LINE = 'rgba(190,190,190,0.3)'

_HIGHLIGHT_COLORS = {
    'wall_internal': ('rgba(30,110,230,0.85)', 'rgba(15,70,160,1.0)'),              
    'wall_internal_estimated': ('rgba(120,180,240,0.75)', 'rgba(60,120,190,1.0)'),  
    'wall_external_confirmed': ('rgba(230,90,30,0.85)', 'rgba(170,60,10,1.0)'),     
    'wall_external_unknown': ('rgba(230,190,190,0.85)', 'rgba(160,50,50,1.0)'),     
}

_CLASS_HIGHLIGHT_PALETTE = {
    'IfcColumn': ('rgba(120,60,170,0.85)', 'rgba(80,30,120,1.0)'),        
    'IfcBeam': ('rgba(150,100,40,0.85)', 'rgba(100,60,10,1.0)'),          
    'IfcMember': ('rgba(150,100,40,0.7)', 'rgba(100,60,10,0.9)'),         
    'IfcSlab': ('rgba(190,150,80,0.85)', 'rgba(140,100,30,1.0)'),         
    'IfcRoof': ('rgba(190,150,80,0.6)', 'rgba(140,100,30,0.8)'),          
    'IfcCovering': ('rgba(90,160,150,0.85)', 'rgba(40,110,100,1.0)'),     
    'IfcCurtainWall': ('rgba(90,150,220,0.85)', 'rgba(40,100,180,1.0)'),  
    'IfcDoor': ('rgba(190,110,60,0.85)', 'rgba(140,70,20,1.0)'),          
    'IfcWindow': ('rgba(110,200,230,0.85)', 'rgba(50,150,190,1.0)'),      
    'IfcRailing': ('rgba(170,170,60,0.85)', 'rgba(120,120,20,1.0)'),      
    'IfcStair': ('rgba(200,90,150,0.85)', 'rgba(150,40,100,1.0)'),        
    'IfcStairFlight': ('rgba(200,90,150,0.7)', 'rgba(150,40,100,0.9)'),
    'IfcRamp': ('rgba(200,90,150,0.6)', 'rgba(150,40,100,0.8)'),
    'IfcRampFlight': ('rgba(200,90,150,0.5)', 'rgba(150,40,100,0.7)'),
    'IfcLightFixture': ('rgba(230,200,20,0.95)', 'rgba(150,130,0,1.0)'),
    'IfcSensor': ('rgba(230,140,20,0.95)', 'rgba(150,90,0,1.0)'),
    'IfcFireSuppressionTerminal': ('rgba(230,60,60,0.95)', 'rgba(150,20,20,1.0)'),
    'IfcAlarm': ('rgba(180,20,180,0.95)', 'rgba(110,0,110,1.0)'),
}


def _get_highlight_color(cat):
    if cat in _HIGHLIGHT_COLORS:
        return _HIGHLIGHT_COLORS[cat]
    if cat in _UNION_CATEGORY_COLORS:
        return _UNION_CATEGORY_COLORS[cat]
    if cat in _CLASS_HIGHLIGHT_PALETTE:
        return _CLASS_HIGHLIGHT_PALETTE[cat]
    h = abs(hash(cat)) % 360
    return (f'hsla({h},65%,55%,0.85)', f'hsla({h},70%,35%,1.0)')


def _category_label(cat):
    _WALL_LABELS = {
        'wall_internal': '내벽(확정)',
        'wall_internal_estimated': '내벽(추정-관계기반)',
        'wall_external_confirmed': '외벽(판정됨)',
        'wall_external_unknown': '외벽(판정불가)',
        'wall_internal_union': '내벽 전체(확정+추정)',
        'wall_external_union': '외벽 전체(판정+판정불가)',
    }
    if cat in _WALL_LABELS:
        return _WALL_LABELS[cat]
    return ite._label(cat)


_EQUIPMENT_COLOR = 'rgba(220,190,20,0.95)'  

_SPACE_FILL = 'rgba(100,180,120,0.16)'
_SPACE_FILL_SELECTED = 'rgba(230,100,60,0.55)'
_SPACE_LINE = 'rgba(60,140,80,0.5)'
_SPACE_LINE_SELECTED = 'rgba(200,60,20,1.0)'

_SPACE_UNMATCHED_FILL = 'rgba(190,190,190,0.16)'
_SPACE_UNMATCHED_LINE = 'rgba(150,150,150,0.4)'

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
_PAIR_PALETTE_FAINT = [
    'rgba(31,119,180,0.12)', 'rgba(255,127,14,0.12)', 'rgba(44,160,44,0.12)',
    'rgba(214,39,40,0.12)', 'rgba(148,103,189,0.12)', 'rgba(140,86,75,0.12)',
    'rgba(227,119,194,0.12)', 'rgba(127,127,127,0.12)', 'rgba(188,189,34,0.12)',
    'rgba(23,190,207,0.12)',
]


def build_floor_category_highlight(ifc_file, plan_data, wall_classification):
    structural_entities = [el['entity'] for el in plan_data['structural']]

    equipment_entities = []
    seen_eq = set()
    for sp in plan_data['spaces']:
        for e in get_space_contained_equipment(ifc_file, sp['entity']):
            if e.GlobalId not in seen_eq:
                seen_eq.add(e.GlobalId)
                equipment_entities.append(e)

    highlight_map = _build_highlight_map(structural_entities, equipment_entities, wall_classification)
    return highlight_map, equipment_entities


UNION_CATEGORY_EXPANSIONS = {
    'wall_internal_union': ('wall_internal', 'wall_internal_estimated'),
    'wall_external_union': ('wall_external_confirmed', 'wall_external_unknown'),
}
_UNION_CATEGORY_COLORS = {
    'wall_internal_union': ('rgba(75,145,235,0.8)', 'rgba(20,80,170,1.0)'),
    'wall_external_union': ('rgba(230,140,110,0.8)', 'rgba(170,70,30,1.0)'),
}


def expand_category(cat):
    return set(UNION_CATEGORY_EXPANSIONS.get(cat, (cat,)))


def get_legend_items():
    wall_categories = ['wall_internal', 'wall_internal_estimated',
                        'wall_external_confirmed', 'wall_external_unknown',
                        'wall_internal_union', 'wall_external_union']
    nonwall_classes = [c for c in ite.ELEMENT_CLASSIFICATION_TARGET_CLASSES
                        if c not in ('IfcWall', 'IfcWallStandardCase')]
    equipment_classes = list(EQUIPMENT_CLASSES)

    items = []
    for cat in wall_categories + nonwall_classes + equipment_classes:
        fill, _line = _get_highlight_color(cat)
        items.append((cat, _category_label(cat), fill))
    return items


def build_pair_labels(a_to_b):
    a_labels, b_labels = {}, {}
    for i, (a_guid, b_guid) in enumerate(a_to_b.items(), start=1):
        a_labels[a_guid] = i
        if b_guid:
            b_labels[b_guid] = i
    return a_labels, b_labels


def build_plan_figure(plan_data, click_grid_spacing=0.5, selected_guid=None,
                       highlight_map=None, equipment_entities=None, pair_labels=None,
                       wall_segments=None, active_categories=None):
    import plotly.graph_objects as go
    fig = go.Figure()
    highlight_map = highlight_map or {}
    pair_labels = pair_labels or {}
    wall_segments = wall_segments or {}
    active_set = set(active_categories) if active_categories is not None else None

    for el in plan_data['structural']:
        cat = highlight_map.get(el['guid'])
        is_wall = el['class'] in ('IfcWall', 'IfcWallStandardCase')
        segment_poly = wall_segments.get(el['guid']) if is_wall else None
        cat_active = cat is not None and (active_set is None or cat in active_set)

        if highlight_map:
            if cat_active:
                if segment_poly is not None:
                    xs_full, ys_full = _polygon_xy_lists(el['polygon'])
                    fig.add_trace(go.Scatter(
                        x=xs_full, y=ys_full, mode='lines', fill='toself',
                        line=dict(width=0.5, color=_FADED_LINE), fillcolor=_FADED_COLOR,
                        hoverinfo='text', text=el.get('hover') or f"{el['class']} {el['name']}".strip(),
                        showlegend=False,
                    ))
                    fill_c, line_c = _get_highlight_color(cat)
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
                fill_c, line_c = _get_highlight_color(cat)
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

    badge_x, badge_y, badge_text, badge_color, badge_line = [], [], [], [], []
    for sp in plan_data['spaces']:
        is_sel = (selected_guid is not None and sp['guid'] == selected_guid)
        pair_no = pair_labels.get(sp['guid'])

        if is_sel:
            fill_c, line_c, line_w = _SPACE_FILL_SELECTED, _SPACE_LINE_SELECTED, 0
        elif pair_no is not None:
            idx = (pair_no - 1) % len(_PAIR_PALETTE)
            fill_c, line_c, line_w = _PAIR_PALETTE_FAINT[idx], _PAIR_PALETTE_LINE[idx], 0
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
        by_class = defaultdict(list)
        for e in equipment_entities:
            cls = e.is_a()
            if active_set is not None and cls not in active_set:
                continue
            by_class[cls].append(e)

        for cls, ents in by_class.items():
            ex, ey, etext = [], [], []
            for e in ents:
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
            if not ex:
                continue
            marker_color, marker_line = _get_highlight_color(cls) if cls in _CLASS_HIGHLIGHT_PALETTE else (_EQUIPMENT_COLOR, 'rgba(120,100,0,1)')
            fig.add_trace(go.Scatter(
                x=ex, y=ey, mode='markers',
                marker=dict(size=11, color=marker_color, symbol='diamond',
                            line=dict(width=1, color=marker_line)),
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