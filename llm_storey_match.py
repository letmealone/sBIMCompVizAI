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
    안내되어 있다. 사용자가 자신의 프로젝트 기준으로 직접 확인한 gemini-3.1-flash-lite
    수치(RPM 15 / TPM 250,000 / RPD 500)를 아래 기본값으로 사용한다.

RPM/TPM/RPD 세 한도에 대한 대응 방식(중요 - 성격이 다름):
  - RPM(분당 요청수): 이 프로세스 안에서 클라이언트측 슬라이딩 윈도우로 실제 방지 가능
    (아래 _throttle_for_rpm). 다만 Streamlit 앱이 여러 프로세스/컨테이너로 스케일아웃되는
    배포라면 프로세스별로 카운터가 나뉘어 정확도가 떨어진다(단일 프로세스 배포 기준의
    최선의 방어).
  - TPM(분당 토큰수): 이 앱의 프롬프트는 층 개수×층당 표시 공간 이름 개수(각각 상한을
    이미 두고 있음)에 비례하므로 250,000 토큰에 걸릴 가능성은 실제로 낮다. 그래도 극단적으로
    층/공간이 많은 IFC를 위해 안전판(문자수 상한 초과시 공간 이름 목록을 더 줄여 재구성)을 둔다.
  - RPD(일일 요청수): API 키/프로젝트 단위로 걸리는 한도라 이 앱만 통제할 수 있는 값이
    아니다(같은 키를 쓰는 모든 사용자의 호출이 합산됨). 앱 차원에서 할 수 있는 것은
    "불필요한 호출을 만들지 않는 것"뿐이며, 이는 이미 (a) 파일쌍당 1회 호출 + (b) 세션
    캐싱으로 재실행시 재호출 방지로 구현되어 있다. RPD 초과 자체를 앱이 막을 수는 없고,
    초과시 명확한 에러 메시지로 안내하는 것까지가 이 모듈의 책임이다.

    #   3. Similarity of the floor's space (room) composition, i.e. both the room COUNT and the
#      room-TYPE names listed after "spaces:" for each storey (e.g. a floor whose spaces are
#      mostly "Office"/"Meeting Room"/"Lobby" should not be matched to a floor whose spaces are
#      "Parking"/"Mechanical Room" - the two are unlikely to be the same physical floor even if
#      names or elevations look superficially similar). Room names may be in different languages
#      or conventions across the two models, so match by likely real-world room TYPE, not exact
#      string equality. Treat this as a supporting signal alongside 1 and 2, useful especially
#      to break ties or catch cases where names/elevations conflict.


"""
import json
import threading
import time


DEFAULT_MODEL = 'gemini-3.1-flash-lite'

# 사용자가 Google AI Studio에서 직접 확인한 gemini-3.1-flash-lite 무료 티어 한도.
# (Google 공식 문서는 고정 수치를 게시하지 않으므로, 실제 배포 후 값이 바뀌었다면
#  아래 상수만 갱신하면 된다.)
DEFAULT_RPM_LIMIT = 15
DEFAULT_TPM_LIMIT = 250_000
DEFAULT_RPD_LIMIT = 500  # 참고용 표시에만 사용 - 이 프로세스가 강제할 수 있는 값이 아님

# RPM 클라이언트측 스로틀링용 전역 상태. 프로세스(=보통 배포된 Streamlit 서버 1개) 안의
# 모든 세션이 공유해야 의미가 있다 - 세션별로 따로 두면 사용자 3명이 각자 1회씩 눌러도
# 실제로는 API 프로젝트 한도(RPM)를 합산해서 초과할 수 있기 때문이다.
_RPM_LOCK = threading.Lock()
_RPM_CALL_TIMES = []  # 최근 호출 시각들(time.monotonic() 기준, 초)


def _throttle_for_rpm(max_rpm=DEFAULT_RPM_LIMIT, safety_margin=2, max_wait_s=75):
    """직전 60초 안에 이미 (max_rpm - safety_margin)회 호출했다면, 여유가 생길 때까지
    대기한다(최대 max_wait_s초). safety_margin은 이 카운터가 완벽하지 않다는 점
    (여러 프로세스 배포시 부정확) 을 감안한 안전 여유분이다."""
    limit = max(1, max_rpm - safety_margin)
    deadline = time.monotonic() + max_wait_s
    while True:
        with _RPM_LOCK:
            now = time.monotonic()
            while _RPM_CALL_TIMES and now - _RPM_CALL_TIMES[0] > 60:
                _RPM_CALL_TIMES.pop(0)
            if len(_RPM_CALL_TIMES) < limit:
                _RPM_CALL_TIMES.append(now)
                return
            wait_s = 60 - (now - _RPM_CALL_TIMES[0]) + 0.5
        if time.monotonic() + wait_s > deadline:
            # 너무 오래 기다려야 하면(다른 세션들이 계속 호출 중) 포기하고 예외로 알린다 -
            # Streamlit 요청을 무한정 블로킹하는 것보다 사용자에게 재시도를 안내하는 편이 낫다.
            raise LlmStoreyMatchError(
                f"다른 세션의 요청이 많아 RPM 한도({max_rpm}/분) 여유가 나지 않습니다. "
                "잠시 후 다시 시도해주세요."
            )
        time.sleep(max(min(wait_s, 5.0), 0.5))


def _estimate_token_count(text):
    """대략적인 토큰수 추정치(공식 카운터가 아님 - countTokens API를 별도로 부르면
    호출 횟수가 늘어나므로 로컬 휴리스틱만 사용). 영문은 대략 4자/토큰, 한글 등 비영문은
    보수적으로 2자/토큰으로 가정해 다소 과대추정되도록(=안전한 쪽으로) 잡는다."""
    ascii_chars = sum(1 for c in text if ord(c) < 128)
    non_ascii_chars = len(text) - ascii_chars
    return ascii_chars / 4 + non_ascii_chars / 2


class LlmStoreyMatchError(RuntimeError):
    """API 키 누락, 패키지 미설치, 호출 실패, 응답 파싱 실패 등을 알리기 위한 예외."""
    pass


def _build_prompt(storeys_a, storeys_b, elevation_hint_mapping, offset_mm,
                   space_summary_a=None, space_summary_b=None, space_name_cap=15):
    """LLM에 보낼 프롬프트 문자열 생성. 토큰 절약을 위해 불필요한 필드는 넣지 않는다.
    space_summary_a/b: {층이름: {'count': int, 'names': [str, ...]}} (get_storey_space_summary 반환값들의 dict).
    None이면 공간 정보 없이(이름+표고만으로) 진행한다.
    space_name_cap: 층 하나당 프롬프트에 나열할 공간 이름 최대 개수 (TPM 안전판에서 축소용)."""
    def _fmt(storeys, space_summary):
        lines = []
        for s in storeys:
            elev = s['Elevation']
            elev_text = f"{elev:.0f}mm" if elev is not None else "unknown"
            info = (space_summary or {}).get(s['Name'])
            if info and info['count'] > 0 and space_name_cap > 0:
                sample = ", ".join(info['names'][:space_name_cap])
                extra = info['count'] - min(space_name_cap, len(info['names']))
                more = f", +{extra} more" if extra > 0 else ""
                space_text = f", spaces: {info['count']} total [{sample}{more}]"
            elif info and info['count'] > 0:
                space_text = f", spaces: {info['count']} total (names omitted)"
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
  3. 층 이름 간 문맥적 유사성과 해당 층의 전체 면적의 유사성을 고려할 것.

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


def _build_prompt_within_tpm_budget(storeys_a, storeys_b, elevation_hint_mapping, offset_mm,
                                     space_summary_a, space_summary_b, tpm_limit):
    """TPM 안전판: 기본(공간이름 15개) 프롬프트가 예상 토큰수 기준으로 TPM 한도의 60%를
    넘으면(출력 토큰/오차 여유분 확보), 공간 이름 표시 개수를 15->6->2->0(개수만 표시)
    순으로 줄여가며 다시 만든다. 지오메트리 재계산이 필요 없는 순수 문자열 작업이라 비용은
    거의 없다. 실제로 이 앱의 프롬프트 크기(층수 x 층당 공간이름 상한)로는 TPM 250,000에
    걸릴 가능성이 낮지만, 극단적으로 층/공간이 많은 IFC를 위한 방어적 장치다."""
    budget = tpm_limit * 0.6
    for cap in (15, 6, 2, 0):
        prompt = _build_prompt(storeys_a, storeys_b, elevation_hint_mapping, offset_mm,
                                space_summary_a=space_summary_a, space_summary_b=space_summary_b,
                                space_name_cap=cap)
        if _estimate_token_count(prompt) <= budget:
            return prompt, cap
    return prompt, cap  # 그래도 넘으면(비정상적으로 층이 많은 경우) 최소 형태로라도 시도


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
                       space_summary_a=None, space_summary_b=None,
                       rpm_limit=DEFAULT_RPM_LIMIT, tpm_limit=DEFAULT_TPM_LIMIT, max_retries=2):
    """Google Gemini API를 '정확히 1회'(요청 자체는, 429 재시도가 있는 경우 최대
    max_retries회까지) 호출해 두 IFC의 층을 이름 문맥+표고 유사도+공간 구성 유사도를
    함께 고려해 매핑한다. 스트리밍 응답을 사용하므로, on_chunk가 주어지면 텍스트가
    도착하는 대로 on_chunk(누적된_텍스트_전체)를 호출해 화면에 실시간으로 보여줄 수 있다.

    RPM/TPM/RPD 세 한도 대응(중요):
      - RPM: 호출 직전 _throttle_for_rpm()으로 클라이언트측에서 실제로 방지한다
        (같은 프로세스 안의 모든 세션이 공유하는 슬라이딩 윈도우).
      - TPM: 프롬프트 예상 토큰수가 tpm_limit의 60%를 넘으면 공간 이름 표시 개수를
        자동으로 줄여 다시 만든다(_build_prompt_within_tpm_budget).
      - RPD: 앱이 강제할 수 없는 값이다(API 키/프로젝트 단위 합산). 초과시 429 에러가
        나면 재시도해도 소용없으므로(다음날 태평양시간 자정까지 대기 필요) 재시도하지
        않고 바로 그 사실을 담은 에러 메시지를 낸다. RPM/TPM성 429는 일시적이므로
        지수 백오프로 최대 max_retries회 재시도한다.

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
    rpm_limit, tpm_limit : int
        이 모델의 실제 분당 요청/토큰 한도(Google AI Studio에서 확인한 값을 그대로 사용).
    max_retries : int
        RPM/TPM성 일시적 429 오류에 대한 최대 재시도 횟수(지수 백오프). RPD(일일) 초과로
        보이는 429는 재시도하지 않는다.
    on_chunk : callable, optional
        스트리밍 청크가 도착할 때마다 on_chunk(누적_텍스트)로 호출된다. 실시간 UI 표시용.
        (파싱 전의 원본 JSON 텍스트가 그대로 전달되므로, 스트림 도중에는 불완전한 JSON일 수 있다.)

    Returns
    -------
    (mapping, detail) : ({A층이름: B층이름 or None}, [매칭 상세 dict, ...])
        detail의 각 원소: {'a_name', 'b_name', 'confidence', 'reason'}
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

    prompt, _used_cap = _build_prompt_within_tpm_budget(
        storeys_a, storeys_b, elevation_hint_mapping, offset_mm,
        space_summary_a, space_summary_b, tpm_limit,
    )
    client = genai.Client(api_key=api_key)

    last_error = None
    for attempt in range(max_retries + 1):
        _throttle_for_rpm(max_rpm=rpm_limit)  # 호출 직전에 RPM 여유 확보(필요시 대기)
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

            full_text = ''.join(accumulated)
            if not full_text:
                raise LlmStoreyMatchError("Gemini 응답이 비어 있습니다.")
            matches = _parse_response_text(full_text)
            return _finalize_matches(matches, storeys_b)

        except LlmStoreyMatchError:
            raise  # 파싱 실패 등 자체 예외는 재시도해도 소용없으므로 그대로 전달
        except Exception as e:
            msg = str(e)
            is_quota_error = ('429' in msg) or ('RESOURCE_EXHAUSTED' in msg) or ('quota' in msg.lower())
            # RPD(일일) 초과로 보이는 단서(문구에 day/daily 등)가 있으면 재시도해도 무의미
            looks_daily = is_quota_error and any(k in msg.lower() for k in ('daily', 'per day', 'rpd'))
            if is_quota_error and not looks_daily and attempt < max_retries:
                if on_chunk is not None:
                    on_chunk(f"[일시적 요청 한도(RPM/TPM) 초과, {attempt + 1}/{max_retries}회 재시도 대기 중...]")
                time.sleep(5 * (attempt + 1))  # 지수적으로 늘려가며 대기
                last_error = e
                continue
            if looks_daily:
                raise LlmStoreyMatchError(
                    "일일 요청 한도(RPD)를 초과한 것으로 보입니다. 이 한도는 API 키/프로젝트 "
                    "단위로 걸리며 이 앱이 강제로 막을 수 없습니다 - 태평양시간 자정에 초기화될 "
                    "때까지 기다리거나, https://aistudio.google.com/rate-limit 에서 실제 한도를 "
                    f"확인하세요. (원본 오류: {msg})"
                ) from e
            if is_quota_error:
                raise LlmStoreyMatchError(
                    f"요청 한도(RPM/TPM) 초과로 재시도했지만 실패했습니다. 잠시 후 다시 "
                    f"시도해주세요. (원본 오류: {msg})"
                ) from e
            raise LlmStoreyMatchError(f"Gemini API 호출 실패: {e}") from e

    raise LlmStoreyMatchError(f"Gemini API 호출이 반복적으로 실패했습니다: {last_error}")
