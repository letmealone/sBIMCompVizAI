# -*- coding: utf-8 -*-
"""
llm_storey_match.py
--------------------
Google Gemini API(google-genai SDK)를 사용해, 두 IFC의 층(IfcBuildingStorey) 이름 간
문맥적(의미적) 연관성과 표고(Elevation) 유사도를 동시에 고려해 자동 매핑한다.

설계 원칙 (무료 할당량 절약이 최우선 요구사항):
  - IFC 파일쌍(A, B)당 API 호출은 정확히 "1회"다. 층 하나하나마다 호출하지 않고,
    양쪽의 모든 층 이름+표고를 프롬프트 하나에 전부 담아 한 번에 전체 매핑을 받는다.
  - 표고 기반 매핑(floorplan_core.match_storeys, API 호출 없음, 결정론적)을 먼저
    계산해 "힌트"로 프롬프트에 함께 제공한다. 이렇게 하면 LLM이 표고 유사도를 별도로
    다시 추론할 필요 없이 이름 문맥과 결합해 최종 판단만 내리면 되므로, 출력 토큰과
    오류 가능성이 줄어든다(=같은 결과를 더 적은 토큰으로).
  - 호출 자체의 캐싱/재호출 방지는 이 모듈의 책임이 아니라 호출하는 쪽(Streamlit 앱)의
    책임이다 - 이 모듈은 "호출 1번 = 함수 1번 실행"이 되도록 순수하게 유지한다.

주의(사실확인, 2026-07 기준):
  - google-generativeai(구 SDK)는 2025-11-30부로 공식 폐지(deprecated)되어, 이 모듈은
    신규 통합 SDK인 `google-genai` 패키지를 사용한다 (`pip install google-genai`).
  - 무료 할당량의 정확한 RPM/RPD 수치는 Google 공식 문서(ai.google.dev/gemini-api/docs/rate-limits)에
    더 이상 고정 테이블로 게시되지 않고, "Google AI Studio에서 프로젝트별로 확인"하도록
    안내되어 있다. 즉, 이 파일 어디에도 "무료 한도는 하루 N회다" 같은 확정적 수치는
    적어두지 않는다(서드파티 블로그마다 수치가 다르고 신뢰도가 낮음) - 실제 한도는
    https://aistudio.google.com/rate-limit 에서 직접 확인 필요.
  - 여러 출처에서 공통적으로 확인되는 사실은 "Flash-Lite 계열 모델이 Flash/Pro보다
    무료 티어 요청 한도가 더 넉넉하다"는 경향뿐이라, 기본 모델을 flash-lite 계열로 둔다.
"""
import json


DEFAULT_MODEL = 'gemini-3.1-flash-lite'


class LlmStoreyMatchError(RuntimeError):
    """API 키 누락, 패키지 미설치, 호출 실패, 응답 파싱 실패 등을 알리기 위한 예외."""
    pass


def _build_prompt(storeys_a, storeys_b, elevation_hint_mapping, offset_mm,
                   space_summary_a=None, space_summary_b=None):
    """LLM에 보낼 프롬프트 문자열 생성. 토큰 절약을 위해 불필요한 필드는 넣지 않는다.
    space_summary_a/b: {층이름: {'count': int, 'names': [str, ...]}} (get_storey_space_summary 반환값들의 dict).
    None이면 공간 정보 없이(이름+표고만으로) 진행한다."""
    def _fmt(storeys, space_summary):
        lines = []
        for s in storeys:
            elev = s['Elevation']
            elev_text = f"{elev:.0f}mm" if elev is not None else "unknown"
            info = (space_summary or {}).get(s['Name'])
            if info and info['count'] > 0:
                sample = ", ".join(info['names'][:15])
                extra = info['count'] - min(15, len(info['names']))
                more = f", +{extra} more" if extra > 0 else ""
                space_text = f", spaces: {info['count']} total [{sample}{more}]"
            else:
                space_text = ", spaces: none/unknown"
            lines.append(f"- \"{s['Name']}\" (elevation={elev_text}{space_text})")
        return "\n".join(lines)

    hint_lines = []
    for a_name, b_name in (elevation_hint_mapping or {}).items():
        hint_lines.append(f"- \"{a_name}\" -> {json.dumps(b_name, ensure_ascii=False)}")
    hint_text = "\n".join(hint_lines) if hint_lines else "(no candidates found)"

    prompt = f"""You are matching building storeys (floors) between two IFC BIM models of the SAME physical building, authored independently by two different modelers (possibly in different languages/conventions, e.g. "1F", "Level 1", "1층", "지하1층", "B1", "GF", "Roof", "PIT", "R/F" may be equivalent concepts even though the strings differ).

Building A storeys:
{_fmt(storeys_a, space_summary_a)}

Building B storeys:
{_fmt(storeys_b, space_summary_b)}

A purely elevation-based statistical estimate (after correcting for a constant vertical
offset of {offset_mm:.0f}mm between the two models' coordinate origins) suggests this
candidate mapping, which you should use as a starting hint, not as ground truth:
{hint_text}

Task: produce the best 1:1 mapping from each Building A storey to a Building B storey,
considering ALL of the following signals together:
  1. Semantic/contextual similarity of the storey names (primary signal when names are
     informative - e.g. both clearly mean "2nd floor" or both clearly mean "basement 1").
  2. Elevation similarity after the offset correction above (secondary/tie-breaking signal,
     and the main signal when names are ambiguous, generic, or uninformative, e.g. "Story_3").
  3. Similarity of the floor's space (room) composition, i.e. both the room COUNT and the
     room-TYPE names listed after "spaces:" for each storey (e.g. a floor whose spaces are
     mostly "Office"/"Meeting Room"/"Lobby" should not be matched to a floor whose spaces are
     "Parking"/"Mechanical Room" - the two are unlikely to be the same physical floor even if
     names or elevations look superficially similar). Room names may be in different languages
     or conventions across the two models, so match by likely real-world room TYPE, not exact
     string equality. Treat this as a supporting signal alongside 1 and 2, useful especially
     to break ties or catch cases where names/elevations conflict.

Rules:
  - Each A storey maps to at most one B storey, and each B storey is used at most once.
  - If no B storey is a reasonable counterpart for an A storey, map it to null.
  - If signals disagree for a candidate pair, weigh all three together and lower the
    confidence, but briefly say why in "reason" (mention which signal(s) drove the decision,
    e.g. "space composition mismatch despite similar elevation").
  - Do not invent storey or space names that are not in the lists above.

Return ONLY a JSON object with this exact shape, no markdown fences, no extra text:
{{
  "matches": [
    {{"a_name": "<exact A name>", "b_name": "<exact B name or null>", "confidence": "high|medium|low", "reason": "<short reason, max 20 words>"}}
  ]
}}
"""
    return prompt


def _parse_response_text(text):
    """모델 응답에서 JSON을 안전하게 파싱. 코드펜스가 붙어 나오는 경우까지 방어적으로 처리."""
    cleaned = text.strip()
    if cleaned.startswith("```"):
        # ```json ... ``` 또는 ``` ... ``` 형태 제거
        cleaned = cleaned.strip("`")
        if cleaned.lower().startswith("json"):
            cleaned = cleaned[4:]
        cleaned = cleaned.strip()
    try:
        data = json.loads(cleaned)
    except json.JSONDecodeError as e:
        raise LlmStoreyMatchError(f"LLM 응답을 JSON으로 파싱하지 못했습니다: {e}\n원본 응답: {text[:500]}")
    if not isinstance(data, dict) or 'matches' not in data:
        raise LlmStoreyMatchError(f"LLM 응답에 'matches' 키가 없습니다. 원본 응답: {text[:500]}")
    return data['matches']


def _finalize_matches(matches, storeys_b):
    """모델이 반환한 matches 리스트를 검증/정리해 (mapping, detail)로 변환하는 공통 로직."""
    valid_b_names = {s['Name'] for s in storeys_b}
    mapping, detail = {}, []
    used_b = set()
    for m in matches:
        a_name = m.get('a_name')
        b_name = m.get('b_name')
        if b_name is not None and (b_name not in valid_b_names or b_name in used_b):
            # 존재하지 않는 이름을 지어냈거나 중복 매핑인 경우 방어적으로 null 처리
            b_name = None
        if b_name is not None:
            used_b.add(b_name)
        mapping[a_name] = b_name
        detail.append({
            'a_name': a_name,
            'b_name': b_name,
            'confidence': m.get('confidence', 'unknown'),
            'reason': m.get('reason', ''),
        })
    return mapping, detail


def match_storeys_llm(storeys_a, storeys_b, offset_mm, api_key,
                       model_name=DEFAULT_MODEL, elevation_hint_mapping=None, on_chunk=None,
                       space_summary_a=None, space_summary_b=None):
    """Google Gemini API를 '정확히 1회' 호출해 두 IFC의 층을 이름 문맥+표고 유사도+공간
    구성 유사도를 함께 고려해 매핑한다. 스트리밍 응답을 사용하므로(요청 자체는 여전히
    1건), on_chunk가 주어지면 텍스트가 도착하는 대로 on_chunk(누적된_텍스트_전체)를
    호출해 화면에 실시간으로 보여줄 수 있다.

    Parameters
    ----------
    storeys_a, storeys_b : floorplan_core.load_ifc()['storeys'] 형태의 리스트
        각 원소는 {'Name': str, 'Elevation': float|None, 'entity': ...} 딕셔너리.
    offset_mm : float
        floorplan_core.match_storeys()가 반환한 좌표계 표고 오프셋(mm). 프롬프트에 사용.
    api_key : str
        Google AI API 키. 비어있으면 즉시 예외.
    model_name : str
        사용할 Gemini 모델명. 무료 할당량이 넉넉한 flash-lite 계열을 기본값으로 사용.
    elevation_hint_mapping : dict, optional
        floorplan_core.match_storeys()가 반환한 {A층이름: B층이름} 매핑(API 호출 없이
        결정론적으로 계산됨). 프롬프트에 힌트로 함께 넣어 LLM이 표고 유사도를 처음부터
        다시 추론하지 않고 이름 문맥과 결합해 판단하게 한다. None이면 힌트 없이 진행.
    space_summary_a, space_summary_b : dict, optional
        {층이름: floorplan_core.get_storey_space_summary() 반환값} 형태. 각 층의 공간
        (Space) 개수와 이름 구성을 프롬프트에 포함시켜, 층 이름/표고가 모호하거나
        서로 충돌할 때 '이 층에 어떤 종류의 공간이 몇 개 있는가'까지 함께 판단 근거로
        쓰도록 한다. 지오메트리 계산이 필요 없는 값이라 추가 비용이 거의 없다.
    on_chunk : callable, optional
        스트리밍 청크가 도착할 때마다 on_chunk(누적_텍스트)로 호출된다. 실시간 UI 표시용.
        (파싱 전의 원본 JSON 텍스트가 그대로 전달되므로, 스트림 도중에는 불완전한 JSON일 수 있다.)

    Returns
    -------
    (mapping, detail) : ({A층이름: B층이름 or None}, [매칭 상세 dict, ...])
        detail의 각 원소: {'a_name', 'b_name', 'confidence', 'reason'}
        API 호출(스트리밍 요청 1건)은 이 함수 안에서 정확히 1회만 발생한다.
    """
    if not api_key:
        raise LlmStoreyMatchError("Google AI API Key가 비어 있습니다.")

    try:
        from google import genai
        from google.genai import types
    except ImportError as e:
        raise LlmStoreyMatchError(
            "google-genai 패키지가 설치되어 있지 않습니다. "
            "`pip install google-genai`로 설치해주세요."
        ) from e

    prompt = _build_prompt(storeys_a, storeys_b, elevation_hint_mapping, offset_mm,
                            space_summary_a=space_summary_a, space_summary_b=space_summary_b)
    client = genai.Client(api_key=api_key)

    accumulated = []
    try:
        stream = client.models.generate_content_stream(
            model=model_name,
            contents=prompt,
            config=types.GenerateContentConfig(
                response_mime_type='application/json',
                temperature=0.0,  # 매핑 작업은 결정론적일수록 좋음 (재현성)
            ),
        )
        for chunk in stream:
            piece = getattr(chunk, 'text', None)
            if piece:
                accumulated.append(piece)
                if on_chunk is not None:
                    on_chunk(''.join(accumulated))
    except Exception as e:
        raise LlmStoreyMatchError(f"Gemini API 호출 실패: {e}") from e

    full_text = ''.join(accumulated)
    if not full_text:
        raise LlmStoreyMatchError("Gemini 응답이 비어 있습니다.")

    matches = _parse_response_text(full_text)
    return _finalize_matches(matches, storeys_b)
