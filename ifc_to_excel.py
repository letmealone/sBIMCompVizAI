# -*- coding: utf-8 -*-
"""
ifc_to_excel.py
----------------
임의의 IFC 파일(.ifc)을 입력받아, 건축공학의 일반적 정보 계층
(Project > Site > Building > Storey > Space > Element)에 맞춰
행렬(엔티티 x 속성) 구조의 엑셀 파일로 자동 추출하는 모듈.
"""

import json
import math
import re
from collections import Counter, OrderedDict, defaultdict

import ifcopenshell
import ifcopenshell.util.element as E
import numpy as np
import pandas as pd
import openpyxl
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

CATEGORY_LABELS = {
    'IfcSite': '대지(Site)',
    'IfcBuilding': '건물(Building)',
    'IfcBuildingStorey': '층(Storey)',
    'IfcSpace': '공간(Space)',
    'IfcWall': '벽(Wall)',
    'IfcWallStandardCase': '벽(Wall)',
    'IfcCurtainWall': '커튼월(CurtainWall)',
    'IfcColumn': '기둥(Column)',
    'IfcBeam': '보(Beam)',
    'IfcMember': '부재(Member)',
    'IfcSlab': '바닥(Slab)',
    'IfcCovering': '천장외장(Covering)',
    'IfcRoof': '지붕(Roof)',
    'IfcRailing': '난간(Railing)',
    'IfcStair': '계단(Stair)',
    'IfcStairFlight': '계단참(StairFlight)',
    'IfcRamp': '경사로(Ramp)',
    'IfcRampFlight': '경사로참(RampFlight)',
    'IfcDoor': '문(Door)',
    'IfcWindow': '창(Window)',
    'IfcOpeningElement': '개구부(Opening)',
    'IfcPlate': '플레이트(Plate)',
    'IfcGrid': '그리드(Grid)',
    'IfcFurnishingElement': '적재물(Furnishing)',
    'IfcFlowTerminal': '설비단말(FlowTerminal)',
    'IfcDistributionElement': '설비요소(Distribution)',
    'IfcLightFixture': '조명(LightFixture)',
    'IfcSensor': '센서(Sensor)',
    'IfcFireSuppressionTerminal': '소방장치(FireSuppressionTerminal)',
    'IfcAlarm': '경보기(Alarm)',
}

CORE_ATTRS = ['GlobalId', 'Name', 'Description', 'ObjectType', 'Tag', 'PredefinedType', 'LongName']

EXCLUDE_FROM_WIDE = {'IfcAnnotation'}
MAX_ROWS_LONG = 300_000  

def _safe_attr(ent, name):
    try:
        v = getattr(ent, name)
    except Exception:
        return None
    if v is None:
        return None
    if hasattr(v, 'is_a'):
        try:
            return v.is_a()
        except Exception:
            return str(v)
    return v


def _storey_via_decomposition(ent):
    cur = ent
    seen = set()
    while cur is not None and hasattr(cur, 'is_a'):
        if cur.is_a('IfcBuildingStorey'):
            return cur
        if id(cur) in seen:
            break
        seen.add(id(cur))
        decomposes = getattr(cur, 'Decomposes', None)
        if not decomposes:
            return None
        cur = decomposes[0].RelatingObject
    return None


def _get_storey(ent):
    try:
        c = E.get_container(ent)
    except Exception:
        c = None
    cur = c
    seen = set()
    while cur is not None and hasattr(cur, 'is_a') and not cur.is_a('IfcBuildingStorey'):
        if id(cur) in seen:
            break
        seen.add(id(cur))
        try:
            cur = E.get_container(cur)
        except Exception:
            cur = None
    if cur is None:
        cur = _storey_via_decomposition(ent)
    return cur


def _get_spaces_via_boundary(ifc_file, ent):
    spaces = _get_spaces_list_via_boundary(ifc_file, ent)
    return ', '.join(spaces) if spaces else None


def _get_spaces_list_via_boundary(ifc_file, ent):
    spaces = []
    for rel in ifc_file.by_type('IfcRelSpaceBoundary'):
        if rel.RelatedBuildingElement == ent and rel.RelatingSpace is not None:
            spaces.append(rel.RelatingSpace.Name)
    return spaces


def _get_material(ent):
    try:
        m = E.get_material(ent)
    except Exception:
        return None
    if m is None:
        return None
    try:
        if m.is_a('IfcMaterial'):
            return m.Name
        if m.is_a('IfcMaterialList'):
            return ', '.join([mm.Name for mm in m.Materials])
        if m.is_a() == 'IfcMaterialLayerSetUsage' and m.ForLayerSet:
            return m.ForLayerSet.LayerSetName or 'LayerSet'
        if hasattr(m, 'Name') and m.Name:
            return m.Name
        return m.is_a()
    except Exception:
        return None


def _flatten_psets(ent):
    try:
        psets = E.get_psets(ent, qtos_only=False)
    except Exception:
        psets = {}
    flat = {}
    for pset_name, props in psets.items():
        for k, v in props.items():
            if k == 'id':
                continue
            flat[f"{pset_name}.{k}"] = v
    return flat


def _resolve_body_items(ent):
    if not ent.Representation:
        return []
    items = []
    for rep in ent.Representation.Representations:
        if rep.RepresentationIdentifier != 'Body':
            continue
        for it in rep.Items:
            if it.is_a('IfcMappedItem'):
                mr = it.MappingSource.MappedRepresentation
                items.extend(mr.Items)
            else:
                items.append(it)
    return items


def _profile_xy_extent(profile):
    try:
        if profile.is_a('IfcRectangleProfileDef'):
            return profile.XDim, profile.YDim
        if profile.is_a('IfcCircleProfileDef'):
            return profile.Radius * 2, profile.Radius * 2
        curve = getattr(profile, 'OuterCurve', None)
        if curve is not None and curve.is_a('IfcIndexedPolyCurve'):
            pts = curve.Points
            if pts.is_a('IfcCartesianPointList2D'):
                arr = np.array(pts.CoordList, dtype=float)
                ext = arr.max(axis=0) - arr.min(axis=0)
                return float(ext[0]), float(ext[1])
    except Exception:
        pass
    return None, None


def _get_local_dimensions(ent):
    if ent.Representation:
        for rep in ent.Representation.Representations:
            if rep.RepresentationIdentifier == 'Box':
                for it in rep.Items:
                    if it.is_a('IfcBoundingBox'):
                        return it.XDim, it.YDim, it.ZDim, 'BoundingBox(원본기록값)'

    items = _resolve_body_items(ent)

    for it in items:
        if it.is_a('IfcExtrudedAreaSolid'):
            x, y = _profile_xy_extent(it.SweptArea)
            if x is not None:
                return x, y, it.Depth, '단면좌표+Depth(직접계산)'

    for it in items:
        if it.is_a('IfcPolygonalFaceSet'):
            coords = it.Coordinates.CoordList
            arr = np.array(coords, dtype=float)
            ext = arr.max(axis=0) - arr.min(axis=0)
            return float(ext[0]), float(ext[1]), float(ext[2]), '메쉬좌표bbox(직접계산)'

    return None, None, None, None


def _dimension_columns(ent, length_unit_scale=0.001):
    x, y, z, src = _get_local_dimensions(ent)
    conv = lambda v: round(v * length_unit_scale, 4) if v is not None else None
    return {
        '치수_X(m)': conv(x),
        '치수_Y(m)': conv(y),
        '치수_Z(m)': conv(z),
        '치수산출방식': src,
    }

DIMENSION_TARGET_CLASSES = {'IfcWall', 'IfcWallStandardCase', 'IfcColumn', 'IfcBeam', 'IfcMember', 'IfcCurtainWall'}


def _polyline_points_2d(curve):
    try:
        if curve.is_a('IfcIndexedPolyCurve'):
            pts = curve.Points
            if pts.is_a('IfcCartesianPointList2D'):
                return [tuple(p) for p in pts.CoordList]
        elif curve.is_a('IfcPolyline'):
            return [tuple(p.Coordinates) for p in curve.Points]
    except Exception:
        pass
    return None


def _shoelace_area(points):
    n = len(points)
    if n < 3:
        return None
    area = 0.0
    for i in range(n):
        x1, y1 = points[i][0], points[i][1]
        x2, y2 = points[(i + 1) % n][0], points[(i + 1) % n][1]
        area += x1 * y2 - x2 * y1
    return abs(area) / 2.0


def _profile_area(profile):
    try:
        if profile.is_a('IfcRectangleProfileDef'):
            return profile.XDim * profile.YDim
        if profile.is_a('IfcCircleProfileDef'):
            return math.pi * profile.Radius ** 2
        outer = getattr(profile, 'OuterCurve', None)
        if outer is None:
            return None
        pts = _polyline_points_2d(outer)
        if not pts:
            return None
        area = _shoelace_area(pts)
        if area is None:
            return None
        if profile.is_a('IfcArbitraryProfileDefWithVoids'):
            for inner in profile.InnerCurves:
                ipts = _polyline_points_2d(inner)
                if ipts:
                    iarea = _shoelace_area(ipts)
                    if iarea:
                        area -= iarea
        return area
    except Exception:
        return None


def _mesh_surface_area(coords, faces_idx):
    arr = np.array(coords, dtype=float)
    total = 0.0
    for idx in faces_idx:
        if len(idx) < 3:
            continue
        p0 = arr[idx[0]]
        for k in range(1, len(idx) - 1):
            p1, p2 = arr[idx[k]], arr[idx[k + 1]]
            total += float(np.linalg.norm(np.cross(p1 - p0, p2 - p0))) / 2.0
    return total


def _get_footprint_area(ent):
    items = _resolve_body_items(ent)

    for it in items:
        if it.is_a('IfcExtrudedAreaSolid'):
            area = _profile_area(it.SweptArea)
            if area is not None:
                return area, '단면좌표(신발끈공식) 직접계산'

    for it in items:
        if it.is_a('IfcPolygonalFaceSet'):
            coords = it.Coordinates.CoordList
            faces_idx = []
            for face in it.Faces:
                ci = face.CoordIndex
                faces_idx.append(tuple(i - 1 for i in ci))
            if faces_idx:
                return _mesh_surface_area(coords, faces_idx), '메쉬삼각형합산 직접계산'
        if it.is_a('IfcTriangulatedFaceSet'):
            coords = it.Coordinates.CoordList
            faces_idx = [tuple(i - 1 for i in tri) for tri in it.CoordIndex]
            if faces_idx:
                return _mesh_surface_area(coords, faces_idx), '메쉬삼각형합산 직접계산'

    return None, None


_AREA_KEY_RE = re.compile(r'area', re.IGNORECASE)
_RATIO_KEY_RE = re.compile(r'ratio', re.IGNORECASE)


def _get_qto_area(flat_props):
    candidates = []
    for key, val in flat_props.items():
        if not isinstance(val, (int, float)):
            continue
        if not key.startswith('Qto_'):
            continue
        if not _AREA_KEY_RE.search(key) or _RATIO_KEY_RE.search(key):
            continue
        score = 0 if 'Gross' in key else (1 if 'Net' in key else 2)
        candidates.append((score, key, val))
    if not candidates:
        return None, None
    candidates.sort(key=lambda c: c[0])
    _, key, val = candidates[0]
    return val, key


def _get_pset_area_fallback(flat_props):
    candidates = []
    for key, val in flat_props.items():
        if not isinstance(val, (int, float)):
            continue
        if not _AREA_KEY_RE.search(key) or _RATIO_KEY_RE.search(key):
            continue
        score = (0 if key.startswith('Qto_') else 1, 0 if 'Gross' in key else 1)
        candidates.append((score, key, val))
    if not candidates:
        return None, None
    candidates.sort(key=lambda c: c[0])
    _, key, val = candidates[0]
    return val, key


SYSTEM_ASSEMBLY_CLASSES = {'IfcDoor', 'IfcWindow', 'IfcCurtainWall'}

def _get_system_bounding_face_area(ent):
    x, y, z, src = _get_local_dimensions(ent)
    dims = [d for d in (x, y, z) if d is not None]
    if len(dims) < 2:
        return None, None
    dims.sort(reverse=True)
    area = dims[0] * dims[1]
    return area, f'시스템 전체 bounding치수 기반(폭x높이, 두께제외; {src})'


def _get_decomposed_children(ent, _depth=0, _max_depth=3):
    if _depth >= _max_depth:
        return []
    children = []
    for rel in (getattr(ent, 'IsDecomposedBy', None) or []):
        for child in rel.RelatedObjects:
            children.append(child)
            children.extend(_get_decomposed_children(child, _depth + 1, _max_depth))
    return children


def _get_assembly_bounding_face_area(ent):
    import ifcopenshell.util.placement as plc

    children = _get_decomposed_children(ent)
    if not children:
        return None, None

    all_corners = []
    for child in children:
        x, y, z, _src = _get_local_dimensions(child)
        if x is None or y is None or z is None:
            continue
        if child.ObjectPlacement is None:
            continue
        try:
            m = plc.get_local_placement(child.ObjectPlacement)
        except Exception:
            continue
        for sx in (0, x):
            for sy in (0, y):
                for sz in (0, z):
                    world_pt = m @ np.array([sx, sy, sz, 1.0])
                    all_corners.append(world_pt[:3])

    if len(all_corners) < 2:
        return None, None
    arr = np.array(all_corners)
    extent = arr.max(axis=0) - arr.min(axis=0)
    dims = sorted(extent, reverse=True)
    area = float(dims[0] * dims[1])
    return area, f'하위부품{len(children)}개 월드bounding 근사(폭x높이, 두께제외)'


def _get_union_footprint_area_m2(ent, tol=0.05):
    try:
        import ifcopenshell.geom as geom
        from shapely.geometry import Polygon
        from shapely.ops import unary_union
    except ImportError:
        return None, None
    
    settings = geom.settings()
    settings.set('use-world-coords', True)
    
    def _get_poly(e):
        try:
            shape = geom.create_shape(settings, e)
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
        
        polys = []
        for rel in getattr(e, 'IsDecomposedBy', []):
            for child in rel.RelatedObjects:
                cp = _get_poly(child)
                if cp is not None and not cp.is_empty:
                    polys.append(cp)
        if polys:
            try:
                u = unary_union(polys)
                if not u.is_empty:
                    return u
            except Exception:
                pass
        return None

    poly = _get_poly(ent)
    if poly is not None and not poly.is_empty:
        return poly.area, '하위부품 평면(footprint) 합집합 직접계산'
    return None, None


def _area_columns(ent, flat_props, length_unit_scale=0.001):
    area, source_key = _get_qto_area(flat_props)
    if area is not None:
        return {'면적(㎡)': round(area, 4), '면적산출방식': f'Qto값 사용({source_key})'}

    if ent.is_a() == 'IfcCurtainWall':
        area_m2, method = _get_union_footprint_area_m2(ent)
        if area_m2 is not None:
            return {'면적(㎡)': round(area_m2, 4), '면적산출방식': method}

    area, method = None, None
    if ent.is_a() in SYSTEM_ASSEMBLY_CLASSES:
        area, method = _get_system_bounding_face_area(ent)
        if area is None:
            area, method = _get_assembly_bounding_face_area(ent)

    if area is None:
        area, method = _get_footprint_area(ent)

    if area is None:
        area, source_key = _get_pset_area_fallback(flat_props)
        if area is not None:
            method = f'Pset값 사용({source_key})'
        else:
            method = '산출불가(좌표계산 실패 + Pset값 없음)'
    else:
        area = area * (length_unit_scale ** 2)
        
    return {
        '면적(㎡)': round(area, 4) if area is not None else None,
        '면적산출방식': method,
    }


AREA_TARGET_CLASSES = {'IfcRoof', 'IfcCovering', 'IfcSlab', 'IfcDoor', 'IfcWindow', 'IfcCurtainWall'}



# ===================================================================
# 3-1b. 벽 내/외벽 판정 및 검증 로직 개선 (선 분할 후 판정 방식)
# ===================================================================

def _determine_wall_classification(ifc_file):
    """
    [수정사항] 선(先) 분할(Segmentation), 후(後) 판정 로직 적용.
    기존에는 벽 전체(IfcWall) 단위로 판정하여 L자 꺾인 벽이 내/외벽 역할을 동시에 
    수행할 때 모순이 발생했음. 이를 해결하기 위해 공간(Space) 단위로 벽체를 가상으로 
    조각(Segment) 낸 뒤, 각 조각이 다른 공간과 교차하는지를 개별 판정함.
    
    반환 형태: 
    {
      (Wall_GlobalId, Space_GlobalId): ('내벽'|'외벽'|'판정불가', '근거'),
      Wall_GlobalId: ('혼합(내/외벽 복합)'|'내벽'|'외벽', '요약 근거') # 엑셀 및 구버전 호환용
    }
    """
    result = {}
    
    try:
        import floorplan_core as fc
    except ImportError:
        fc = None

    if not fc:
        return {}

    # 1. 공간별 바닥 폴리곤(Footprint) 캐싱
    space_fps = {}
    storey_spaces = defaultdict(list)
    for sp in ifc_file.by_type('IfcSpace'):
        fp = fc.get_footprint_polygon(sp)
        if fp and not fp.is_empty:
            space_fps[sp.GlobalId] = fp
            st = _get_storey(sp)
            if st:
                storey_spaces[st.GlobalId].append((sp.GlobalId, fp))

    # 2. 벽-공간 관계 매핑
    wall_space_map = defaultdict(list)
    for r in ifc_file.by_type('IfcRelSpaceBoundary'):
        w = r.RelatedBuildingElement
        s = r.RelatingSpace
        if w and w.is_a('IfcWall') and s:
            wall_space_map[w.GlobalId].append(s)

    # 3. 조각(Segment) 단위 교차 검사
    for w_guid, spaces in wall_space_map.items():
        w_ent = ifc_file.by_id(w_guid)
        w_fp = fc.get_footprint_polygon(w_ent)
        
        if not w_fp or w_fp.is_empty:
            for s in spaces:
                result[(w_guid, s.GlobalId)] = ('판정불가', '벽체 지오메트리 추출 실패')
            continue

        for s in spaces:
            s_fp = space_fps.get(s.GlobalId)
            if not s_fp:
                result[(w_guid, s.GlobalId)] = ('판정불가', '공간 지오메트리 추출 실패')
                continue
            
            # [핵심 로직] 공간 영역에 맞게 벽체를 분할(Segmentation)하여 해당 공간과 맞닿은 조각 획득
            seg_result = fc.get_space_wall_segment_polygon(ifc_file, w_ent, s, w_fp)
            if not seg_result or not seg_result[0]:
                result[(w_guid, s.GlobalId)] = ('외벽(추정)', '벽체 분할(Segmentation) 실패로 기본 외벽 간주')
                continue
            
            seg_poly, _method = seg_result
            
            # 분할된 벽체 조각이 반대편에 있는 다른 공간과 닿아있는지 확인 (두께나 모델링 이격 극복을 위해 10cm 버퍼)
            buffered_seg = seg_poly.buffer(0.1)
            
            st = _get_storey(s)
            other_spaces = storey_spaces.get(st.GlobalId, []) if st else []
            
            is_internal = False
            for other_guid, other_fp in other_spaces:
                if other_guid == s.GlobalId:
                    continue
                
                # 다른 공간과 교차하는 영역이 유의미한지 검사
                if buffered_seg.intersects(other_fp):
                    inter = buffered_seg.intersection(other_fp)
                    if inter.area > 0.01:  # 단순 선 접촉이 아닌 실제 면적 공유
                        is_internal = True
                        break
                        
            if is_internal:
                result[(w_guid, s.GlobalId)] = ('내벽', '조각 단위 교차 검사: 반대편에 다른 공간 존재함')
            else:
                result[(w_guid, s.GlobalId)] = ('외벽', '조각 단위 교차 검사: 반대편이 외부와 접함')

    # 4. 엑셀 출력 및 하위 호환성을 위해 단일 GlobalId에 대한 대표 속성(혼합 포함)도 병기
    for w_guid, spaces in wall_space_map.items():
        internal_cnt = sum(1 for s in spaces if result.get((w_guid, s.GlobalId), ('', ''))[0] == '내벽')
        external_cnt = sum(1 for s in spaces if result.get((w_guid, s.GlobalId), ('', ''))[0] in ('외벽', '외벽(추정)'))
        
        if internal_cnt > 0 and external_cnt > 0:
            result[w_guid] = ('혼합(내/외벽 복합)', '분할 조각 중 내벽과 외벽 속성 혼재')
        elif internal_cnt > 0:
            result[w_guid] = ('내벽', '모든 분할 조각이 내벽으로 판정됨')
        elif external_cnt > 0:
            result[w_guid] = ('외벽', '모든 분할 조각이 외벽으로 판정됨')
        else:
            result[w_guid] = ('판정불가', '근거 없음')
            
    return result


# ===================================================================
# 3-1c. 벽 외 모든 구조부재 대상 범용 내/외부 판정
# ===================================================================

ELEMENT_CLASSIFICATION_TARGET_CLASSES = (
    'IfcSlab', 'IfcRoof', 'IfcCovering',
    'IfcColumn', 'IfcBeam', 'IfcMember', 'IfcCurtainWall', 'IfcDoor', 'IfcWindow',
    'IfcRailing', 'IfcStair', 'IfcStairFlight', 'IfcRamp', 'IfcRampFlight',
)

def _element_both_sides_space_check(ifc_file, target_classes=None):
    import ifcopenshell.util.placement as plc

    normals = {}
    for r in ifc_file.by_type('IfcRelSpaceBoundary'):
        elem = r.RelatedBuildingElement
        if elem is None:
            continue
        if target_classes is not None and elem.is_a() not in target_classes:
            continue
        space = r.RelatingSpace
        cg = r.ConnectionGeometry
        if space is None or cg is None:
            continue
        surf = cg.SurfaceOnRelatingElement
        if surf is None or not surf.is_a('IfcCurveBoundedPlane'):
            continue
        plane = surf.BasisSurface
        try:
            space_m = plc.get_local_placement(space.ObjectPlacement)
            plane_m = plc.get_axis2placement(plane.Position)
        except Exception:
            continue
        world_m = space_m @ plane_m
        normal = world_m[:3, 2]
        nrm = np.linalg.norm(normal)
        if nrm == 0:
            continue
        normals.setdefault(elem.GlobalId, []).append(normal / nrm)

    result = {}
    for gid, ns in normals.items():
        ref = ns[0]
        has_same = any(np.dot(n, ref) > 0.5 for n in normals)
        has_opp = any(np.dot(n, ref) < -0.5 for n in normals)
        result[gid] = has_same and has_opp
    return result


def _element_distinct_space_count(ifc_file, target_classes=None):
    counts = defaultdict(set)
    for r in ifc_file.by_type('IfcRelSpaceBoundary'):
        elem = r.RelatedBuildingElement
        if elem is None or r.RelatingSpace is None:
            continue
        if target_classes is not None and elem.is_a() not in target_classes:
            continue
        counts[elem.GlobalId].add(r.RelatingSpace.GlobalId)
    return {gid: len(spaces) for gid, spaces in counts.items()}


def _determine_element_classification(ifc_file, target_classes=ELEMENT_CLASSIFICATION_TARGET_CLASSES):
    both_sides = _element_both_sides_space_check(ifc_file, target_classes)
    distinct_count = _element_distinct_space_count(ifc_file, target_classes)

    result = {}
    for cls in target_classes:
        for ent in ifc_file.by_type(cls):
            gid = ent.GlobalId
            if gid in result:
                continue 
            
            if gid in both_sides:
                if both_sides[gid]:
                    result[gid] = ('내부', '2차: RelSpaceBoundary 양면 Space 접촉 확인')
                else:
                    result[gid] = ('외부(추정)', '2차: 한쪽 면만 Space 접촉 → 외부로 추정')
            elif distinct_count.get(gid, 0) >= 2:
                n = distinct_count[gid]
                result[gid] = ('내부(추정-관계기반)', f'3차: 서로 다른 Space {n}개와 연결')
            else:
                result[gid] = ('판정불가', '1차/2차/3차 모두 근거 데이터 없음')
    return result


# ===================================================================
# 3-2. 공간(Space) - 부재(Element) 1:1 매칭 시트
# ===================================================================

def _build_space_element_matrix(ifc_file, wall_classification=None):
    if wall_classification is None:
        wall_classification = _determine_wall_classification(ifc_file)
    rows = []
    for rel in ifc_file.by_type('IfcRelSpaceBoundary'):
        sp = rel.RelatingSpace
        elem = rel.RelatedBuildingElement
        if sp is None or elem is None:
            continue
        storey = _storey_via_decomposition(sp) or _get_storey(sp)
        
        # [수정] 엑셀 내보내기 시 (Wall_GUID, Space_GUID) 조합의 세부 판정 결과 기록
        if elem.is_a('IfcWall'):
            cls_result, cls_reason = wall_classification.get((elem.GlobalId, sp.GlobalId), 
                                     wall_classification.get(elem.GlobalId, (None, None)))
        else:
            cls_result, cls_reason = None, None
            
        rows.append({
            '층(Storey)': storey.Name if storey else None,
            '공간(Space)_Name': sp.Name,
            '공간(Space)_GUID': sp.GlobalId,
            '부재_IFC_Class': elem.is_a(),
            '부재_Name': _safe_attr(elem, 'Name'),
            '부재_GUID': elem.GlobalId,
            'PhysicalOrVirtual': rel.PhysicalOrVirtualBoundary,
            'InternalOrExternal': rel.InternalOrExternalBoundary,
            '벽_내외벽_판정(조각단위)': cls_result,
            '벽_판정근거': cls_reason,
            'RelSpaceBoundary_GUID': rel.GlobalId,
        })
    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.sort_values(['층(Storey)', '공간(Space)_Name', '부재_IFC_Class', '부재_Name']).reset_index(drop=True)
    return df


def _build_wall_classification_sheet(ifc_file, wall_classification=None):
    if wall_classification is None:
        wall_classification = _determine_wall_classification(ifc_file)
    rows = []
    for w in ifc_file.by_type('IfcWall'):
        storey = _get_storey(w)
        # [수정] 혼합 여부가 저장된 단일 GlobalId 키를 불러와 객체(Wall) 기준의 대푯값 기록
        result, reason = wall_classification.get(w.GlobalId, ('판정불가', '근거 데이터 없음'))
        rows.append({
            '층(Storey)': storey.Name if storey else None,
            'Name': _safe_attr(w, 'Name'),
            'GlobalId': w.GlobalId,
            '내외벽_판정(객체단위)': result,
            '판정근거': reason,
        })
    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.sort_values(['층(Storey)', '내외벽_판정(객체단위)', 'Name']).reset_index(drop=True)
    return df


# ===================================================================
# 3-3. 공간(Space)에 매칭되지 않은 부재 시트
# ===================================================================

def _build_relspaceboundary_duplication_report(ifc_file):
    pair_count = defaultdict(int)
    elem_class = {}
    for r in ifc_file.by_type('IfcRelSpaceBoundary'):
        elem = r.RelatedBuildingElement
        sp = r.RelatingSpace
        if elem is None or sp is None:
            continue
        pair_count[(sp.GlobalId, elem.GlobalId)] += 1
        elem_class[elem.GlobalId] = elem.is_a()

    raw_by_class = defaultdict(int)
    unique_pairs_by_class = defaultdict(int)
    for (sp_guid, elem_guid), cnt in pair_count.items():
        cls = elem_class[elem_guid]
        raw_by_class[cls] += cnt
        unique_pairs_by_class[cls] += 1

    rows = []
    for cls in sorted(set(raw_by_class) | set(unique_pairs_by_class)):
        raw = raw_by_class[cls]
        uniq = unique_pairs_by_class[cls]
        rows.append({
            'IFC_Class': cls,
            '원시_관계건수': raw,
            '고유_공간-부재쌍_수': uniq,
            '초과분(같은공간에중복연결)': raw - uniq,
            '중복여부': '있음' if raw > uniq else '없음',
        })
    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.sort_values('초과분(같은공간에중복연결)', ascending=False).reset_index(drop=True)
    return df


def _build_unmatched_elements(ifc_file):
    rsb = ifc_file.by_type('IfcRelSpaceBoundary')
    matched_guids = {rel.RelatedBuildingElement.GlobalId
                      for rel in rsb if rel.RelatedBuildingElement is not None}
    matched_classes = {rel.RelatedBuildingElement.is_a()
                        for rel in rsb if rel.RelatedBuildingElement is not None}

    space_storeys = set()
    for sp in ifc_file.by_type('IfcSpace'):
        st = _storey_via_decomposition(sp) or _get_storey(sp)
        if st:
            space_storeys.add(st.Name)

    rows = []
    for ent in ifc_file.by_type('IfcElement'):
        if ent.GlobalId in matched_guids:
            continue
        cls = ent.is_a()
        storey = _get_storey(ent)
        storey_name = storey.Name if storey else None

        if cls not in matched_classes:
            reason = '①구조상 비대상(이 모델에서 해당 클래스는 RelSpaceBoundary 미사용)'
        elif storey_name not in space_storeys:
            reason = '②해당 층에 Space 없음'
        else:
            reason = '③매칭누락 의심(같은 층에 Space 있는데 비어있음)'

        rows.append({
            '층(Storey)': storey_name,
            'IFC_Class': cls,
            'Name': _safe_attr(ent, 'Name'),
            'GlobalId': ent.GlobalId,
            'PredefinedType': _safe_attr(ent, 'PredefinedType'),
            '미매칭_사유': reason,
        })
    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.sort_values(['미매칭_사유', '층(Storey)', 'IFC_Class', 'Name']).reset_index(drop=True)
    return df


def _label(cls):
    return CATEGORY_LABELS.get(cls, cls)


# ===================================================================
# 1. 계층 구조 추출 (Project > Site > Building > Storey > Space)
# ===================================================================

def _build_hierarchy(ifc_file):
    rows = []

    def add(level, ifc_class, guid, name, parent_path, extra=''):
        rows.append({
            'Level': level, 'IFC_Class': ifc_class, 'GlobalId': guid,
            'Name': name, '상위경로': parent_path, '비고': extra,
        })

    projects = ifc_file.by_type('IfcProject')
    proj = projects[0] if projects else None
    if proj:
        add(0, 'IfcProject', proj.GlobalId, proj.Name, '')
    proj_name = proj.Name if proj else ''

    sites = ifc_file.by_type('IfcSite')
    if not sites:
        sites = [None]

    for site in sites:
        if site is None:
            buildings = ifc_file.by_type('IfcBuilding')
            path0 = proj_name
        else:
            add(1, 'IfcSite', site.GlobalId, site.Name, proj_name)
            path0 = f"{proj_name} > {site.Name}"
            buildings = []
            for rel in (site.IsDecomposedBy or []):
                buildings += [o for o in rel.RelatedObjects if o.is_a('IfcBuilding')]
            if not buildings:
                buildings = ifc_file.by_type('IfcBuilding')

        for building in buildings:
            path1 = f"{path0} > {building.Name}"
            add(2, 'IfcBuilding', building.GlobalId, building.Name, path0)

            storeys = []
            for rel in (building.IsDecomposedBy or []):
                storeys += [o for o in rel.RelatedObjects if o.is_a('IfcBuildingStorey')]
            storeys.sort(key=lambda s: (s.Elevation if s.Elevation is not None else 0))

            for st in storeys:
                path2 = f"{path1} > {st.Name}"
                add(3, 'IfcBuildingStorey', st.GlobalId, st.Name, path1, extra=f"Elevation={st.Elevation}")

                contained = []
                for rel in (st.ContainsElements or []):
                    contained += list(rel.RelatedElements)
                spaces = [c for c in contained if c.is_a('IfcSpace')]
                others = [c for c in contained if not c.is_a('IfcSpace')]

                for sp in spaces:
                    add(4, 'IfcSpace', sp.GlobalId, sp.Name, path2, extra=(sp.LongName or ''))
                if others:
                    add(4, '(요소 합계)', '', f"{len(others)}개 요소 직접 포함", path2,
                        extra=', '.join(sorted(set(o.is_a() for o in others))))

    return pd.DataFrame(rows)


# ===================================================================
# 2. 전체 속성 Long 포맷 (모든 IfcProduct x 모든 Attribute/Pset/Qto)
# ===================================================================

def _build_long_format(ifc_file):
    rows = []
    products = ifc_file.by_type('IfcProduct')

    for ent in products:
        cls = ent.is_a()
        storey = _get_storey(ent)
        storey_name = storey.Name if storey else None
        space_name = None if cls == 'IfcSpace' else _get_spaces_via_boundary(ifc_file, ent)
        material = _get_material(ent)
        guid = _safe_attr(ent, 'GlobalId')
        name = _safe_attr(ent, 'Name')

        for attr in CORE_ATTRS:
            v = _safe_attr(ent, attr)
            if v is None:
                continue
            rows.append({
                'IFC_Class': cls, 'GlobalId': guid, 'Name': name,
                '층(Storey)': storey_name, '공간(Space)': space_name, '재질(Material)': material,
                '구분': 'Attribute', 'Pset': '-', '속성명': attr, '값': v,
            })

        for k, v in _flatten_psets(ent).items():
            pset_name, prop_name = k.split('.', 1)
            kind = 'Qto' if pset_name.startswith('Qto_') else 'Pset'
            rows.append({
                'IFC_Class': cls, 'GlobalId': guid, 'Name': name,
                '층(Storey)': storey_name, '공간(Space)': space_name, '재질(Material)': material,
                '구분': kind, 'Pset': pset_name, '속성명': prop_name, '값': v,
            })

        if len(rows) > MAX_ROWS_LONG:
            break

    return pd.DataFrame(rows)


# ===================================================================
# 3. 클래스별 Wide(행렬) 포맷 - IFC 안에 실제 존재하는 클래스만 동적 생성
# ===================================================================

def _build_wide_sheets(ifc_file):
    products = ifc_file.by_type('IfcProduct')
    by_class = OrderedDict()
    for ent in products:
        cls = ent.is_a()
        if cls in EXCLUDE_FROM_WIDE:
            continue
        by_class.setdefault(cls, []).append(ent)

    priority = ['IfcSite', 'IfcBuilding', 'IfcBuildingStorey', 'IfcSpace']
    ordered_classes = [c for c in priority if c in by_class]
    ordered_classes += sorted(
        [c for c in by_class if c not in priority],
        key=lambda c: -len(by_class[c])
    )

    sheets = OrderedDict()
    for cls in ordered_classes:
        ents = by_class[cls]
        rows = []
        for ent in ents:
            storey = _get_storey(ent)
            flat = _flatten_psets(ent)
            row = {
                'GlobalId': _safe_attr(ent, 'GlobalId'),
                'Name': _safe_attr(ent, 'Name'),
                'ObjectType': _safe_attr(ent, 'ObjectType'),
                'PredefinedType': _safe_attr(ent, 'PredefinedType'),
                'Tag': _safe_attr(ent, 'Tag'),
                '층(Storey)': storey.Name if storey else None,
                '공간(Space)': None if cls == 'IfcSpace' else _get_spaces_via_boundary(ifc_file, ent),
                '재질(Material)': _get_material(ent),
            }
            if cls in DIMENSION_TARGET_CLASSES:
                row.update(_dimension_columns(ent))
            if cls in AREA_TARGET_CLASSES:
                row.update(_area_columns(ent, flat))
            row.update(flat)
            rows.append(row)
        sheets[_label(cls)] = pd.DataFrame(rows)
    return sheets


def load_interest_set(spec_xlsx_path):
    wb = openpyxl.load_workbook(spec_xlsx_path, data_only=True)
    interest = set()

    if '속성목록' in wb.sheetnames:
        ws = wb['속성목록']
        header = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
        try:
            i_ns   = header.index('Namespace') + 1
            i_prop = header.index('속성명_IFC') + 1
        except ValueError:
            i_ns, i_prop = None, None

        if i_ns and i_prop:
            for r in range(2, ws.max_row + 1):
                ns   = ws.cell(row=r, column=i_ns).value
                prop = ws.cell(row=r, column=i_prop).value
                if ns and prop:
                    interest.add((str(ns).strip(), str(prop).strip()))
            return interest

    def _parse_token(token):
        token = token.strip()
        if not token or any(x in token for x in ['->', '(', ')']):
            return None
        if token.startswith('IfcRel') and '.' not in token:
            return None
        parts = token.split('.')
        if len(parts) < 2:
            return None
        ns, prop = (parts[-2].strip(), parts[-1].strip()) if len(parts) >= 3 \
                   else (parts[0].strip(), parts[1].strip())
        return (ns, prop) if ns and prop else None

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        header_row, col_idx = None, None
        for r in range(1, min(ws.max_row + 1, 15)):
            for c in range(1, ws.max_column + 1):
                v = ws.cell(row=r, column=c).value
                if isinstance(v, str) and 'IFC' in v and '스키마' in v:
                    header_row, col_idx = r, c
                    break
            if header_row:
                break
        if header_row is None:
            continue
        for r in range(header_row + 1, ws.max_row + 1):
            raw = ws.cell(row=r, column=col_idx).value
            if not raw or not isinstance(raw, str):
                continue
            for eq_part in raw.split('='):
                for slash_part in eq_part.split('/'):
                    result = _parse_token(slash_part.strip())
                    if result:
                        interest.add(result)

    return interest


def _format_props_cell(flat_dict, interest_set=None):
    if not flat_dict:
        return ''

    grouped = OrderedDict()
    for k, v in flat_dict.items():
        pset_name, prop_name = k.split('.', 1)
        grouped.setdefault(pset_name, []).append((prop_name, v))

    if not interest_set:
        lines = []
        for pset_name, props in grouped.items():
            lines.append(f"[{pset_name}]")
            for prop_name, v in props:
                v_str = '' if v is None else str(v)
                lines.append(f"  {prop_name} = {v_str}")
        return '\n'.join(lines)

    black = InlineFont(color='000000')
    red = InlineFont(color='FF0000', b=True)
    blocks = []
    buf = ''  

    def flush_black():
        nonlocal buf
        if buf:
            blocks.append(TextBlock(black, buf))
            buf = ''

    for pset_name, props in grouped.items():
        buf += f"[{pset_name}]\n"
        for prop_name, v in props:
            v_str = '' if v is None else str(v)
            line = f"  {prop_name} = {v_str}\n"
            is_interest = (pset_name, prop_name) in interest_set
            if is_interest:
                flush_black()
                blocks.append(TextBlock(red, line))
            else:
                buf += line
    flush_black()

    if not blocks:
        return ''
    last = blocks[-1]
    if last.text.endswith('\n'):
        blocks[-1] = TextBlock(last.font, last.text[:-1])
    return CellRichText(*blocks)


# ===================================================================
# 3-1. 단일 셀(All-in-one cell) 통합 시트
# ===================================================================

def _build_consolidated_sheet(ifc_file, interest_set=None):
    products = ifc_file.by_type('IfcProduct')
    rows = []
    for ent in products:
        cls = ent.is_a()
        if cls in EXCLUDE_FROM_WIDE:
            continue
        storey = _get_storey(ent)
        flat = _flatten_psets(ent)
        base = {
            'IFC_Class': cls,
            '분류': _label(cls),
            '층(Storey)': storey.Name if storey else None,
            'Name': _safe_attr(ent, 'Name'),
            'GlobalId': _safe_attr(ent, 'GlobalId'),
            'ObjectType': _safe_attr(ent, 'ObjectType'),
            'PredefinedType': _safe_attr(ent, 'PredefinedType'),
            'Tag': _safe_attr(ent, 'Tag'),
            '재질(Material)': _get_material(ent),
        }
        if cls in DIMENSION_TARGET_CLASSES:
            base.update(_dimension_columns(ent))
        if cls in AREA_TARGET_CLASSES:
            base.update(_area_columns(ent, flat))
        base['속성개수'] = len(flat)
        base['전체속성(Pset.속성 = 값)'] = _format_props_cell(flat, interest_set)

        space_list = [] if cls == 'IfcSpace' else _get_spaces_list_via_boundary(ifc_file, ent)
        if not space_list:
            row = dict(base)
            row['공간(Space)'] = None
            rows.append(row)
        else:
            for sp_name in space_list:
                row = dict(base)
                row['공간(Space)'] = sp_name
                rows.append(row)

    df = pd.DataFrame(rows)
    if not df.empty:
        priority_order = [
            '층(Storey)', '공간(Space)', 'IFC_Class', '분류', 'Name',
            '치수_X(m)', '치수_Y(m)', '치수_Z(m)', '면적(㎡)', '재질(Material)',
            '속성개수', '전체속성(Pset.속성 = 값)', '치수산출방식', '면적산출방식',
        ]
        col_order = [c for c in priority_order if c in df.columns]
        col_order += [c for c in df.columns if c not in col_order]  
        df = df[col_order]
    return df


# ===================================================================
# 4. 엑셀 작성 (서식 포함)
# ===================================================================

HEADER_FILL = PatternFill('solid', start_color='1F4E78', end_color='1F4E78')
HEADER_FONT = Font(bold=True, color='FFFFFF', name='맑은 고딕', size=10)
BODY_FONT = Font(name='맑은 고딕', size=10)
THIN = Side(style='thin', color='D9D9D9')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _write_df(ws, df):
    if df.empty:
        ws.append(['(데이터 없음)'])
        return
    ws.append(list(df.columns))
    for c in range(1, len(df.columns) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = BORDER
    for _, r in df.iterrows():
        vals = []
        for v in r.tolist():
            if isinstance(v, (list, tuple)) and not isinstance(v, CellRichText):
                v = ', '.join(str(x) for x in v)
            vals.append(v)
        ws.append(vals)
    for rr in range(2, ws.max_row + 1):
        max_lines = 1
        for cc in range(1, len(df.columns) + 1):
            cell = ws.cell(row=rr, column=cc)
            cell.font = BODY_FONT
            cell.border = BORDER
            n_lines = str(cell.value).count('\n') + 1 if cell.value else 1
            if n_lines > 1:
                cell.alignment = Alignment(vertical='top', wrap_text=True)
                max_lines = max(max_lines, n_lines)
            else:
                cell.alignment = Alignment(vertical='center')
        if max_lines > 1:
            ws.row_dimensions[rr].height = min(max_lines * 14, 400)
    sample = df.iloc[:200] if len(df) > 200 else df
    for c in range(1, len(df.columns) + 1):
        col_name = str(df.columns[c - 1])
        if '전체속성' in col_name or '속성정보' in col_name:
            ws.column_dimensions[get_column_letter(c)].width = 80
            continue
        maxlen = max(
            [len(str(df.columns[c - 1]))] +
            [len(str(v).split('\n')[0]) for v in sample.iloc[:, c - 1].astype(str).tolist()]
        )
        ws.column_dimensions[get_column_letter(c)].width = min(max(maxlen + 2, 10), 45)
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions


def _unique_sheet_name(wb, name, used):
    base = name[:31]
    candidate = base
    i = 2
    while candidate in used:
        suffix = f"_{i}"
        candidate = base[: 31 - len(suffix)] + suffix
        i += 1
    used.add(candidate)
    return candidate


# ===================================================================
# 메인 함수
# ===================================================================

def extract_ifc_to_excel(ifc_path: str, output_path: str = None, include_long: bool = True,
                          consolidated: bool = True, per_class_wide: bool = True,
                          space_element_matrix: bool = True, spec_path: str = None,
                          status_cb=None) -> str:
    if output_path is None:
        output_path = ifc_path.rsplit('.', 1)[0] + '_추출.xlsx'

    ifc_file = ifcopenshell.open(ifc_path)

    interest_set = load_interest_set(spec_path) if spec_path else None

    cnt = Counter(e.is_a() for e in ifc_file)
    df_summary = pd.DataFrame(sorted(cnt.items(), key=lambda x: -x[1]), columns=['IFC_Class', '개수'])
    df_summary.loc[len(df_summary)] = ['IFC Schema', ifc_file.schema]
    if interest_set is not None:
        df_summary.loc[len(df_summary)] = ['관심속성(spec) 매칭개수', len(interest_set)]

    df_hierarchy = _build_hierarchy(ifc_file)

    wb = Workbook()
    wb.remove(wb.active)
    used_names = set()

    ws = wb.create_sheet(_unique_sheet_name(wb, '00_요약', used_names))
    _write_df(ws, df_summary)

    ws = wb.create_sheet(_unique_sheet_name(wb, '01_계층구조', used_names))
    _write_df(ws, df_hierarchy)

    if consolidated:
        df_consol = _build_consolidated_sheet(ifc_file, interest_set)
        ws = wb.create_sheet(_unique_sheet_name(wb, '02_통합조회(단일셀)', used_names))
        _write_df(ws, df_consol)

    if include_long:
        df_long = _build_long_format(ifc_file)
        ws = wb.create_sheet(_unique_sheet_name(wb, '03_전체속성(Long)', used_names))
        _write_df(ws, df_long)

    if space_element_matrix:
        wall_classification = _determine_wall_classification(ifc_file)

        df_matrix = _build_space_element_matrix(ifc_file, wall_classification=wall_classification)
        ws = wb.create_sheet(_unique_sheet_name(wb, '04_공간-부재_매칭(1대1)', used_names))
        _write_df(ws, df_matrix)

        df_unmatched = _build_unmatched_elements(ifc_file)
        ws = wb.create_sheet(_unique_sheet_name(wb, '05_미매칭부재(Space없음)', used_names))
        _write_df(ws, df_unmatched)

        df_wall_class = _build_wall_classification_sheet(ifc_file, wall_classification=wall_classification)
        ws = wb.create_sheet(_unique_sheet_name(wb, '06_벽_내외벽_판정', used_names))
        _write_df(ws, df_wall_class)

        df_dup = _build_relspaceboundary_duplication_report(ifc_file)
        ws = wb.create_sheet(_unique_sheet_name(wb, '07_RelSpaceBoundary_중복진단', used_names))
        _write_df(ws, df_dup)

    if per_class_wide:
        wide_sheets = _build_wide_sheets(ifc_file)
        for sheet_label, df in wide_sheets.items():
            ws = wb.create_sheet(_unique_sheet_name(wb, sheet_label, used_names))
            _write_df(ws, df)

    wb.save(output_path)
    return output_path