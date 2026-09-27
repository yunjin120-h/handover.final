"""
ai_engine.py — 업무 인수인계 시스템의 'AI 두뇌' 모듈 (OpenAI API)

하는 일
  1. 업로드 파일(xlsx/csv/docx/pdf/txt)을 텍스트로 읽기
  2. 엑셀 표 구조 자동 인식 → 기존 코드가 쓰는 '표준 열 이름'으로 변환
  3. 근거 기반 Q&A (근거 행 번호는 AI가 아니라 Python이 계산)
  4. 회의록·캘린더·메일에서 일정/마감/메일 추출 → 온보딩 일과표
  5. 인수인계서 초안(JSON) 생성 → 기존 검토 화면·DOCX 생성 코드 그대로 사용

AI를 쓸 수 없으면(키 없음, 네트워크 오류 등) 호출하는 쪽에서
기존 규칙 기반 코드(memo_parser.py, qa_engine.py)로 자동 전환합니다.
"""

import io
import json
import os
import re
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import streamlit as st

# ---------------------------------------------------------------
# 설정
# ---------------------------------------------------------------
DEFAULT_MODEL = "gpt-5.6-terra"
# 비용 표시용 (gpt-5.6-terra 기준, 100만 토큰당 달러). 실제 청구액은 OpenAI 사이트에서 확인하세요.
PRICE_INPUT_PER_M = 2.0
PRICE_OUTPUT_PER_M = 12.0
KRW_PER_USD = 1400

# 기존 대시보드·일과표·Q&A 코드가 참조하는 표준 열 구조
CANONICAL = {
    "schedule": {
        "label": "업무일정",
        "columns": ["일자", "업무", "프로젝트/현장", "담당", "상태", "우선순위",
                    "진행률", "다음 조치", "마감", "관련 파일", "비고"],
    },
    "project": {
        "label": "프로젝트 진행현황",
        "columns": ["프로젝트/현장", "현재 단계", "진행률", "최근 완료", "다음 액션",
                    "의사결정/확인 필요", "리스크/이슈", "마감일", "외부 담당자",
                    "내부 협업부서", "관련 문서"],
    },
    "contact": {
        "label": "담당자 연락망",
        "columns": ["구분", "회사/부서", "성명", "직책", "연락 수단", "연락처",
                    "관련 업무", "커뮤니케이션 유의사항", "대체 연락자"],
    },
    "asset": {
        "label": "계정·권한·자산",
        "columns": ["시스템/자산", "유형", "현재 사용자", "후임자", "권한 수준",
                    "인계 방법", "상태", "완료 목표일", "주의사항"],
    },
}

# Q&A 근거 표시용 기본 파일/시트 이름 (규칙 기반 모드와 동일)
DEFAULT_SOURCE = {
    "schedule": ("S", "01_업무일정.xlsx", "업무일정"),
    "project": ("P", "02_프로젝트_진행현황.xlsx", "프로젝트현황"),
    "contact": ("C", "03_담당자_연락망.xlsx", "담당자연락망"),
    "asset": ("A", "04_계정_권한_자산.xlsx", "계정권한자산"),
}

COMMON_RULES = (
    "너는 한국 회사의 업무 인수인계 자료를 분석하는 도우미다.\n"
    "- 반드시 제공된 자료에 실제로 있는 내용만 사용한다. 자료에 없는 사실, 이름, 날짜, 숫자를 지어내지 않는다.\n"
    "- 확인되지 않는 값은 빈 문자열로 둔다.\n"
    "- 날짜는 YYYY-MM-DD, 시간은 HH:MM(24시간) 형식으로 쓴다.\n"
    "- 날짜 옆 요일 표기가 실제 날짜와 맞지 않으면 날짜 숫자를 따른다.\n"
    "- 모든 설명은 한국어로 쓴다.\n"
)


# ---------------------------------------------------------------
# API 키 / 상태
# ---------------------------------------------------------------
def _read_secrets_file():
    candidates = [
        Path.cwd() / ".streamlit" / "secrets.toml",
        Path(__file__).resolve().parent / ".streamlit" / "secrets.toml",
    ]
    for path in candidates:
        if path.exists():
            try:
                import tomllib
                return tomllib.loads(path.read_text(encoding="utf-8"))
            except Exception:
                return {}
    return {}


def _setting(name, default=""):
    value = os.environ.get(name, "").strip()
    if value:
        return value
    value = _read_secrets_file().get(name)
    if value:
        return str(value).strip()
    try:
        if name in st.secrets:
            return str(st.secrets[name]).strip()
    except Exception:
        pass
    return default


def get_model():
    return _setting("OPENAI_MODEL", DEFAULT_MODEL)


def ai_status():
    """(사용 가능 여부, 사용 불가 이유)"""
    try:
        import openai  # noqa: F401
    except ImportError:
        return False, "openai 라이브러리가 없어요. 터미널에서 `pip install openai` 를 실행하세요."

    key = _setting("OPENAI_API_KEY")
    if not key:
        return False, "API 키가 없어요. `.streamlit/secrets.toml` 파일에 OPENAI_API_KEY 를 넣어주세요."
    if not key.isascii() or not key.startswith("sk-"):
        return False, "API 키 형식이 이상해요. sk- 로 시작하는 키 전체를 따옴표 안에 붙여넣었는지 확인하세요."
    return True, ""


def ai_available():
    return ai_status()[0]


def _add_usage(input_tokens, output_tokens):
    try:
        usage = st.session_state.setdefault(
            "ai_usage", {"calls": 0, "input": 0, "output": 0}
        )
        usage["calls"] += 1
        usage["input"] += int(input_tokens or 0)
        usage["output"] += int(output_tokens or 0)
    except Exception:
        pass


def usage_summary():
    try:
        usage = st.session_state.get("ai_usage") or {}
    except Exception:
        usage = {}
    calls = usage.get("calls", 0)
    tokens_in = usage.get("input", 0)
    tokens_out = usage.get("output", 0)
    usd = tokens_in / 1e6 * PRICE_INPUT_PER_M + tokens_out / 1e6 * PRICE_OUTPUT_PER_M
    return {
        "calls": calls,
        "input": tokens_in,
        "output": tokens_out,
        "krw": int(round(usd * KRW_PER_USD)),
    }


class AIError(Exception):
    pass


def _call_tool(system, user_text, tool_name, schema, max_tokens=4096):
    """OpenAI를 호출하고, 정해진 JSON 형식(schema)으로만 답을 받는다."""
    ok, reason = ai_status()
    if not ok:
        raise AIError(reason)

    from openai import OpenAI

    client = OpenAI(api_key=_setting("OPENAI_API_KEY"), timeout=300.0, max_retries=2)
    request = dict(
        model=get_model(),
        max_completion_tokens=max(max_tokens * 3, 8000),
        messages=[
            {"role": "system", "content": system},
            {"role": "user", "content": user_text},
        ],
        tools=[{
            "type": "function",
            "function": {
                "name": tool_name,
                "description": "분석 결과를 반드시 이 형식으로 반환한다.",
                "parameters": schema,
            },
        }],
        tool_choice={"type": "function", "function": {"name": tool_name}},
    )
    # GPT-5.6 계열은 이 방식(함수 도구)을 쓸 때 reasoning_effort 를 'none' 으로 둬야 한다
    effort = _setting("OPENAI_REASONING_EFFORT", "none")
    if effort:
        request["reasoning_effort"] = effort

    try:
        response = client.chat.completions.create(**request)
    except Exception as exc:
        # 모델이 reasoning_effort 설정 자체를 모르는 경우(구형 모델 등) 빼고 한 번 더 시도
        if "reasoning_effort" in str(exc) and "reasoning_effort" in request:
            request.pop("reasoning_effort")
            response = client.chat.completions.create(**request)
        else:
            raise

    usage = getattr(response, "usage", None)
    if usage is not None:
        _add_usage(getattr(usage, "prompt_tokens", 0), getattr(usage, "completion_tokens", 0))

    choice = response.choices[0]
    if choice.finish_reason == "length":
        raise AIError("AI 응답이 너무 길어서 중간에 잘렸어요. 자료 양을 줄여서 다시 시도해주세요.")

    message = choice.message
    for call in getattr(message, "tool_calls", None) or []:
        if call.function.name == tool_name:
            try:
                return json.loads(call.function.arguments)
            except json.JSONDecodeError as exc:
                raise AIError("AI 응답 형식이 깨져 있어요. 다시 시도해주세요.") from exc

    # 혹시 도구 호출 대신 본문에 JSON을 쓴 경우
    content = (getattr(message, "content", "") or "").strip()
    match = re.search(r"\{.*\}", content, re.DOTALL)
    if match:
        try:
            return json.loads(match.group(0))
        except json.JSONDecodeError:
            pass
    raise AIError("AI가 정해진 형식으로 답하지 않았어요. 다시 시도해주세요.")


# ---------------------------------------------------------------
# 공통 유틸
# ---------------------------------------------------------------
def cell_text(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    if isinstance(value, (pd.Timestamp, datetime, date)):
        ts = pd.Timestamp(value)
        if ts.hour == 0 and ts.minute == 0:
            return ts.strftime("%Y-%m-%d")
        return ts.strftime("%Y-%m-%d %H:%M")
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip().replace("\r", " ").replace("\n", " ")


def _decode(data):
    for encoding in ("utf-8-sig", "cp949", "euc-kr"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="ignore")


def _norm_date(value):
    text = str(value or "").strip()
    match = re.search(r"(\d{4})[-./](\d{1,2})[-./](\d{1,2})", text)
    if not match:
        return ""
    try:
        return date(int(match.group(1)), int(match.group(2)), int(match.group(3))).isoformat()
    except ValueError:
        return ""


def _norm_time(value):
    match = re.search(r"(\d{1,2}):(\d{2})", str(value or ""))
    if not match:
        return ""
    return f"{int(match.group(1)):02d}:{match.group(2)}"


def _s(value):
    """AI 결과 값을 안전한 문자열로."""
    if value is None:
        return ""
    if isinstance(value, (list, dict)):
        return json.dumps(value, ensure_ascii=False)
    return str(value).strip()


def _obj(keys):
    return {
        "type": "object",
        "properties": {k: {"type": "string"} for k in keys},
        "required": list(keys),
    }


def _arr(keys):
    return {"type": "array", "items": _obj(keys)}


# ---------------------------------------------------------------
# 1. 파일 → 텍스트
# ---------------------------------------------------------------
def _docx_text(data):
    from docx import Document
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    document = Document(io.BytesIO(data))
    lines = []
    for child in document.element.body.iterchildren():
        tag = child.tag.split("}")[-1]
        if tag == "p":
            text = Paragraph(child, document).text.strip()
            if text:
                lines.append(text)
        elif tag == "tbl":
            for row in Table(child, document).rows:
                cells = []
                for cell in row.cells:
                    text = cell.text.strip().replace("\n", " ")
                    if not cells or cells[-1] != text:  # 병합 셀 중복 제거
                        cells.append(text)
                if any(cells):
                    lines.append(" | ".join(cells))
    return "\n".join(lines)


def _pdf_text(data):
    try:
        from pypdf import PdfReader
    except ImportError as exc:
        raise ValueError("PDF를 읽으려면 터미널에서 `pip install pypdf` 를 실행하세요.") from exc
    reader = PdfReader(io.BytesIO(data))
    return "\n".join((page.extract_text() or "") for page in reader.pages)


def read_raw_sheets(name, data):
    """헤더 없이 원본 그대로 읽기. {시트명: DataFrame} (행 번호 = index + 1)"""
    if Path(name).suffix.lower() == ".csv":
        return {"CSV": pd.read_csv(io.StringIO(_decode(data)), header=None, dtype=object)}
    return pd.read_excel(io.BytesIO(data), sheet_name=None, header=None)


@st.cache_data(show_spinner=False)
def extract_text(name, data):
    ext = Path(name).suffix.lower()
    if ext in (".txt", ".md"):
        return _decode(data)
    if ext == ".docx":
        return _docx_text(data)
    if ext == ".pdf":
        return _pdf_text(data)
    if ext in (".xlsx", ".xlsm", ".csv"):
        parts = []
        for sheet, raw in read_raw_sheets(name, data).items():
            parts.append(f"[시트: {sheet}]")
            for idx, row in raw.iterrows():
                cells = [cell_text(v) for v in row.tolist()]
                while cells and not cells[-1]:
                    cells.pop()
                if any(cells):
                    parts.append(f"R{idx + 1}: " + " | ".join(cells))
        return "\n".join(parts)
    raise ValueError(f"지원하지 않는 파일 형식이에요: {name}")


# ---------------------------------------------------------------
# 2. 엑셀 표 구조 자동 인식
# ---------------------------------------------------------------
def ensure_canonical_columns(df, table_type):
    """표준 열이 없으면 빈 열로 추가 (대시보드 KeyError 방지)."""
    if df is None:
        return None
    df = df.copy()
    for column in CANONICAL[table_type]["columns"]:
        if column not in df.columns:
            df[column] = ""
    df.attrs = dict(getattr(df, "attrs", {}))
    return df


def _sheet_preview(raw, max_rows=8):
    lines = []
    for idx, row in raw.head(max_rows + 4).iterrows():
        cells = [cell_text(v)[:40] for v in row.tolist()]
        while cells and not cells[-1]:
            cells.pop()
        if any(cells):
            lines.append(f"R{idx + 1}: " + " | ".join(cells))
        if len(lines) >= max_rows:
            break
    return "\n".join(lines)


def _mapping_system():
    table_desc = []
    for key, spec in CANONICAL.items():
        table_desc.append(f"- {key} ({spec['label']}): " + ", ".join(spec["columns"]))
    return (
        COMMON_RULES
        + "\n엑셀 시트 미리보기를 보고, 각 시트가 아래 표준 표 중 무엇인지와 열 대응 관계를 알려준다.\n"
        + "\n".join(table_desc)
        + "\n\n규칙:\n"
        "- table_type 은 schedule / project / contact / asset / unknown 중 하나.\n"
        "  schedule=날짜별 할 일, project=현장/프로젝트별 진행 현황, contact=사람 연락처, asset=시스템·계정·권한·장비 인계 목록.\n"
        "- header_row 는 실제 열 이름이 적힌 행 번호(R 뒤 숫자). 제목 행은 헤더가 아니다.\n"
        "- column_map 의 source 는 헤더 행에 적힌 원본 열 이름을 글자 그대로, target 은 위 표준 열 이름 중 하나.\n"
        "- 의미가 확실히 같은 열만 연결한다. 확실하지 않으면 넣지 않는다. 하나의 target 에는 source 하나만.\n"
        "- file 과 sheet 는 미리보기 제목에 적힌 값을 그대로 쓴다.\n"
    )


MAPPING_SCHEMA = {
    "type": "object",
    "properties": {
        "sheets": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "file": {"type": "string"},
                    "sheet": {"type": "string"},
                    "table_type": {
                        "type": "string",
                        "enum": ["schedule", "project", "contact", "asset", "unknown"],
                    },
                    "header_row": {"type": "integer"},
                    "column_map": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "source": {"type": "string"},
                                "target": {"type": "string"},
                            },
                            "required": ["source", "target"],
                        },
                    },
                },
                "required": ["file", "sheet", "table_type", "header_row", "column_map"],
            },
        }
    },
    "required": ["sheets"],
}


@st.cache_data(show_spinner=False)
def _ai_map_sheets(preview_text):
    return _call_tool(_mapping_system(), preview_text, "report_table_mapping", MAPPING_SCHEMA, 3000)


def _unique_columns(names):
    seen = {}
    result = []
    for i, name in enumerate(names):
        name = name or f"_col{i}"
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 0
        result.append(name)
    return result


def _apply_mapping(raw, item, table_type):
    columns = CANONICAL[table_type]["columns"]
    column_map = {}
    for entry in item.get("column_map", []) or []:
        source = str(entry.get("source", "")).strip()
        target = str(entry.get("target", "")).strip()
        if source and target in columns and target not in column_map.values():
            column_map[source] = target
    if not column_map:
        return None, None

    def header_at(row_index):
        if 0 <= row_index < len(raw):
            return [cell_text(v) for v in raw.iloc[row_index].tolist()]
        return []

    def hits(row_index):
        header = header_at(row_index)
        return sum(1 for source in column_map if source in header)

    header_index = int(item.get("header_row") or 1) - 1
    if hits(header_index) < max(1, len(column_map) // 2):
        best = max(range(min(len(raw), 15)), key=hits, default=header_index)
        if hits(best) > 0:
            header_index = best

    body = raw.iloc[header_index + 1:].copy()
    body.columns = _unique_columns(header_at(header_index))
    body = body.dropna(how="all")
    body = body[[c for c in body.columns if not c.startswith("_col")]]
    body = body.rename(columns={s: t for s, t in column_map.items() if s in body.columns})

    for column in columns:
        if column not in body.columns:
            body[column] = ""
    extra = [c for c in body.columns if c not in columns]
    body = body[columns + extra]

    # index 를 '헤더 다음 행 = 0' 기준으로 맞추면 엑셀 행 번호 = 헤더행 + 1 + index
    body.index = body.index - (header_index + 1)
    body.attrs["excel_header_row"] = header_index + 1

    info = {
        "인식 결과": CANONICAL[table_type]["label"],
        "헤더 행": header_index + 1,
        "열 매핑": ", ".join(f"{s} → {t}" for s, t in column_map.items() if s != t) or "표준 열 이름과 동일",
    }
    return body, info


def normalize_excel_files(files):
    """
    files: [(파일명, bytes), ...]
    반환: {"schedule": df|None, "project": ..., "contact": ..., "asset": ...,
           "unknown": [파일명], "mapping": [인식 정보]}
    """
    raws = {}
    blocks = []
    for name, data in files:
        for sheet, raw in read_raw_sheets(name, data).items():
            if raw.dropna(how="all").empty:
                continue
            raws[(name, str(sheet))] = raw
            blocks.append(f"### 파일: {name} / 시트: {sheet}\n{_sheet_preview(raw)}")

    result = {key: None for key in CANONICAL}
    result["unknown"] = []
    result["mapping"] = []
    if not blocks:
        return result

    mapping = _ai_map_sheets("\n\n".join(blocks))
    recognized = set()

    for item in mapping.get("sheets", []) or []:
        key = (str(item.get("file", "")).strip(), str(item.get("sheet", "")).strip())
        table_type = item.get("table_type")
        if key not in raws or table_type not in CANONICAL:
            continue
        df, info = _apply_mapping(raws[key], item, table_type)
        if df is None or df.empty:
            continue
        df.attrs["source_file"] = key[0]
        df.attrs["sheet_name"] = key[1]
        df.attrs["table_type"] = table_type
        info = {"파일": key[0], "시트": key[1], **info}
        current = result[table_type]
        if current is None or len(df) > len(current):
            result[table_type] = df
        recognized.add(key[0])
        result["mapping"].append(info)

    result["unknown"] = sorted({name for name, _ in raws} - recognized)
    return result


# ---------------------------------------------------------------
# 3. 근거 기반 Q&A
# ---------------------------------------------------------------
QA_SCHEMA = {
    "type": "object",
    "properties": {
        "answer": {"type": "string"},
        "sources": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "id": {"type": "string"},
                    "evidence": {"type": "string"},
                },
                "required": ["id", "evidence"],
            },
        },
    },
    "required": ["answer", "sources"],
}


def _qa_system(reference_date):
    return (
        COMMON_RULES
        + f"\n너는 업무를 새로 인수받은 후임자의 질문에 답한다. 오늘(기준일)은 {reference_date} 로 간주한다.\n"
        "자료의 각 행/문서 앞에는 [S4], [P5], [D2] 같은 ID가 붙어 있다.\n\n"
        "답변 규칙:\n"
        "- 자료에서 찾은 사실만으로 답한다. 없으면 '업로드된 자료에서 확인되지 않습니다'라고 말하고 어떤 정보가 있으면 답할 수 있는지 한 줄로 안내한다.\n"
        "- 핵심부터 2~6문장 또는 짧은 목록으로 답한다. 날짜·진행률·담당자·다음 조치 같은 핵심 값은 **굵게** 표시한다.\n"
        "- 조건에 맞는 항목이 여러 개면 빠짐없이 모두 나열한다.\n"
        "- '가장', '최대', 'N% 이상', 'N일까지' 같은 질문은 관련 행 전체의 값을 비교해서 정확히 답한다. 날짜 비교는 날짜 값으로 한다.\n"
        "- '이번 주', '오늘' 같은 표현은 기준일을 기준으로 해석한다.\n"
        "- 여러 표를 연결해야 하면 연결한다 (예: 프로젝트의 협업부서 + 연락망의 담당자).\n"
        "- 답변 본문에는 [S4] 같은 ID를 쓰지 않는다.\n"
        "- sources 에는 답변에 실제로 사용한 행/문서 ID를 모두 넣고(최대 10개), evidence 에는 그 행에서 사용한 핵심 내용을 20자 안팎으로 요약한다.\n"
    )


def _build_context(dfs, docs):
    lines = []
    lookup = {}
    for table_type, df in dfs.items():
        if df is None or df.empty:
            continue
        prefix, default_file, default_sheet = DEFAULT_SOURCE[table_type]
        file_name = df.attrs.get("source_file", default_file)
        sheet_name = df.attrs.get("sheet_name", default_sheet)
        header_row = int(df.attrs.get("excel_header_row", 1))
        lines.append(f"\n## 표: {CANONICAL[table_type]['label']} (파일 {file_name})")
        for idx, row in df.iterrows():
            try:
                row_number = header_row + 1 + int(idx)
            except Exception:
                continue
            parts = [f"{col}={cell_text(val)}" for col, val in row.items() if cell_text(val)]
            if not parts:
                continue
            row_id = f"{prefix}{row_number}"
            lines.append(f"[{row_id}] " + " | ".join(parts))
            summary = " / ".join(
                cell_text(row.get(col)) for col in CANONICAL[table_type]["columns"][:3]
                if cell_text(row.get(col))
            )
            lookup[row_id] = {"file": file_name, "sheet": sheet_name, "row": row_number, "evidence": summary}

    for i, (name, text) in enumerate(docs or [], start=1):
        row_id = f"D{i}"
        lines.append(f"\n## 문서 [{row_id}] {name}\n{text[:6000]}")
        lookup[row_id] = {"file": name, "sheet": "문서", "row": None, "evidence": ""}

    return "\n".join(lines), lookup


@st.cache_data(show_spinner=False, max_entries=300)
def _ai_answer_cached(question, context, reference_date):
    user_text = f"# 자료\n{context}\n\n# 질문\n{question}"
    return _call_tool(_qa_system(reference_date), user_text, "answer_with_sources", QA_SCHEMA, 2000)


def default_reference_date(schedule_df=None):
    """Q&A 기준일 기본값: 업무일정의 가장 이른 날짜 (없으면 오늘)."""
    if schedule_df is not None and not schedule_df.empty:
        for column in ["일자", "마감"]:
            if column in schedule_df.columns:
                dates = pd.to_datetime(schedule_df[column], errors="coerce").dropna()
                if not dates.empty:
                    return dates.min().date()
    return date.today()


def answer_question(question, schedule_df=None, project_df=None, contact_df=None,
                    asset_df=None, docs=None, reference_date=None, use_ai=True):
    """
    기존 qa_engine.answer_question 과 같은 형태로 반환:
      {"answer": str, "sources": [{file, sheet, row, evidence}], "engine": "ai"|"rule", ...}
    """
    question = (question or "").strip()
    if not question:
        return {"answer": "질문을 입력해주세요.", "sources": [], "engine": "none"}

    reason = ""
    if use_ai and ai_available():
        try:
            dfs = {"schedule": schedule_df, "project": project_df,
                   "contact": contact_df, "asset": asset_df}
            context, lookup = _build_context(dfs, docs)
            ref = str(reference_date or date.today().isoformat())
            output = _ai_answer_cached(question, context, ref)

            sources = []
            seen = set()
            for item in output.get("sources", []) or []:
                row_id = str(item.get("id", "")).strip().strip("[]")
                if row_id in lookup and row_id not in seen:
                    seen.add(row_id)
                    source = dict(lookup[row_id])
                    evidence = _s(item.get("evidence"))
                    if evidence:
                        source["evidence"] = evidence
                    sources.append(source)

            answer = _s(output.get("answer")) or "답변을 만들지 못했습니다. 질문을 조금 바꿔서 다시 해주세요."
            return {"answer": answer, "sources": sources, "engine": "ai"}
        except Exception as exc:
            reason = f"{type(exc).__name__}: {exc}"
    elif use_ai:
        reason = ai_status()[1]

    from qa_engine import answer_question as rule_answer

    result = rule_answer(
        question,
        schedule_df=schedule_df,
        project_df=project_df,
        contact_df=contact_df,
        asset_df=asset_df,
    )
    result["engine"] = "rule"
    result["fallback_reason"] = reason
    return result


def judge_answer(question, expected, expected_source, answer, sources_text):
    """평가 스크립트용: 기대 답변과 실제 답변 비교."""
    schema = {
        "type": "object",
        "properties": {
            "verdict": {"type": "string", "enum": ["통과", "부분", "실패"]},
            "reason": {"type": "string"},
        },
        "required": ["verdict", "reason"],
    }
    system = (
        "너는 Q&A 시스템 채점자다. 기대 답변의 핵심 사실(대상, 날짜, 숫자, 다음 조치, 인물)이 실제 답변에 모두 맞게 들어 있고 "
        "근거 위치가 기대 근거와 대체로 맞으면 '통과', 핵심은 맞지만 일부 누락·근거 불일치가 있으면 '부분', "
        "핵심 사실이 틀리거나 빠졌으면 '실패'로 판정한다. 기대 답변에 없는 정확한 추가 정보는 감점하지 않는다. "
        "reason 은 한국어 한 문장."
    )
    user_text = (
        f"질문: {question}\n\n기대 답변: {expected}\n기대 근거: {expected_source}\n\n"
        f"실제 답변: {answer}\n실제 근거: {sources_text}"
    )
    return _call_tool(system, user_text, "report_verdict", schema, 500)


# ---------------------------------------------------------------
# 4. 회의록·캘린더·메일 → 온보딩 일정
# ---------------------------------------------------------------
SCHEDULE_SCHEMA = {
    "type": "object",
    "properties": {
        "handover_start_date": {"type": "string"},
        "timed_events": _arr(["date", "start", "end", "title", "detail", "source"]),
        "all_day_items": _arr(["date", "title", "priority", "detail", "source"]),
        "mails": _arr(["date", "time", "subject", "to", "summary", "source"]),
    },
    "required": ["handover_start_date", "timed_events", "all_day_items", "mails"],
}

SCHEDULE_SYSTEM = (
    COMMON_RULES
    + "\n회의록·캘린더·메일 문서에서 후임자 온보딩 일정표에 넣을 항목을 뽑는다.\n\n"
    "- timed_events: 시작·종료 시간이 있는 회의/일정. 회의록이면 '일시'의 날짜와 시간을 쓴다.\n"
    "  title 은 일정 이름, detail 은 참석자·장소·핵심 논의를 한 줄로.\n"
    "- all_day_items: 시간 없이 날짜만 있는 마감·할 일. 캘린더의 '종일/종일 마감' 항목, "
    "회의록 결정사항·Next Action 중 구체적인 날짜가 적힌 것.\n"
    "  마감이면 title 앞에 '[마감] ', 회의록의 조치 항목이면 '[조치] ' 를 붙인다. 기한이 '~9/10'처럼 적혀 있으면 그 날짜를 쓴다.\n"
    "  priority 는 긴급/상/중/하 중 하나 (마감은 보통 상, 조치는 보통 중). detail 은 담당자와 해야 할 일을 한 줄로.\n"
    "- mails: 메일 문서마다 1건. date/time 은 보낸 일시, to 는 받는 사람 이름과 직책, summary 는 요청/안내 핵심 한 줄.\n"
    "- 같은 일정이 캘린더와 회의록에 모두 있으면 한 번만 넣고 source 에 두 파일명을 ' / ' 로 함께 적는다.\n"
    "- source 는 근거 문서의 파일명.\n"
    "- handover_start_date: 후임자가 실제로 업무 인수(1주차)를 시작하는 날. 문서에 '다음 주부터', '1주차' 같은 근거가 있을 때만 "
    "그 주의 월요일 날짜로 쓰고, 근거가 약하면 빈 문자열.\n"
    "- 날짜가 확인되지 않는 항목은 넣지 않는다.\n"
)


@st.cache_data(show_spinner=False)
def _ai_schedule_cached(document_blob):
    return _call_tool(SCHEDULE_SYSTEM, document_blob, "report_schedule_items", SCHEDULE_SCHEMA, 8000)


def extract_schedule_items(docs):
    """docs: [(파일명, 텍스트)] → 일과표에 주입할 데이터"""
    blob = "\n\n".join(f"### 문서: {name}\n{text[:8000]}" for name, text in docs)
    output = _ai_schedule_cached(blob)

    timed = []
    for item in output.get("timed_events", []) or []:
        day, start = _norm_date(item.get("date")), _norm_time(item.get("start"))
        if not day or not start:
            continue
        timed.append({
            "date": day, "start": start, "end": _norm_time(item.get("end")) or start,
            "title": _s(item.get("title")), "detail": _s(item.get("detail")),
            "source": _s(item.get("source")) or "AI 문서 분석",
        })

    all_day = []
    for item in output.get("all_day_items", []) or []:
        day = _norm_date(item.get("date"))
        if not day or not _s(item.get("title")):
            continue
        priority = _s(item.get("priority"))
        all_day.append({
            "date": day, "title": _s(item.get("title")),
            "priority": priority if priority in ("긴급", "상", "중", "하") else "중",
            "detail": _s(item.get("detail")),
            "source": _s(item.get("source")) or "AI 문서 분석",
        })

    mails = []
    for item in output.get("mails", []) or []:
        if not _s(item.get("subject")):
            continue
        mails.append({
            "date": _norm_date(item.get("date")), "time": _norm_time(item.get("time")),
            "subject": _s(item.get("subject")), "to": _s(item.get("to")),
            "summary": _s(item.get("summary")),
            "source": _s(item.get("source")) or "메일",
        })

    return {
        "start_date": _norm_date(output.get("handover_start_date")),
        "timed_events": timed,
        "all_day_items": all_day,
        "mails": mails,
    }


# ---------------------------------------------------------------
# 5. 인수인계서 초안 생성
# ---------------------------------------------------------------
TASK_COLUMNS = ["주요 업무", "업무 목적", "업무 프로세스", "우선순위", "비고"]
DETAIL_FIELDS = [
    "업무명", "업무 개요", "목적 / 성과지표", "진행 중 프로젝트 현황", "정기 업무",
    "비정기 업무", "주요 일정 / 마감", "관련 시스템 / 계정 / 권한", "관련 담당자 / 연락처",
    "협업 부서", "특이사항 / 주의사항", "리스크 / 미해결 이슈",
    "참고 파일 경로 / 문서 링크", "후임자 숙지 필요사항", "인수인계 완료 여부",
]
URGENT_COLUMNS = ["일자", "내용", "대응 방법", "담당"]
MONTHLY_COLUMNS = ["일자", "일정 내용", "비고"]
ASSET_COLUMNS = ["시스템 / 자산명", "유형", "권한 수준", "인계 방법", "상태", "비고"]
CHECK_COLUMNS = ["체크", "항목", "확인일", "확인자", "비고"]
META_KEYS = ["기관명", "부서명", "문서번호", "보존기간"]
BASIC_KEYS = ["소속 부서", "직위 / 직책", "인계자 성명", "인수자 성명", "작성일", "인수인계 완료 예정일"]
SIGN_KEYS = ["인계자 성명", "인계자 서명일", "인수자 성명", "인수자 서명일", "확인자 성명", "확인자 서명일"]

HANDOVER_SCHEMA = {
    "type": "object",
    "properties": {
        "meta": _obj(META_KEYS),
        "basic": _obj(BASIC_KEYS),
        "tasks": _arr(TASK_COLUMNS),
        "details": _arr(DETAIL_FIELDS),
        "urgent_schedule": _arr(URGENT_COLUMNS),
        "monthly_schedule": _arr(MONTHLY_COLUMNS),
        "communication_note": {"type": "string"},
        "assets": _arr(ASSET_COLUMNS),
        "checklist": _arr(CHECK_COLUMNS),
        "signatures": _obj(SIGN_KEYS),
    },
    "required": ["meta", "basic", "tasks", "details", "urgent_schedule", "monthly_schedule",
                 "communication_note", "assets", "checklist", "signatures"],
}


def _handover_system(mode):
    common = (
        COMMON_RULES
        + "\n업무 인수인계서 양식을 채운다. 각 칸의 의미:\n"
        "- meta: 기관명(회사명), 부서명, 문서번호, 보존기간\n"
        "- basic: 소속 부서, 직위 / 직책(인계자 기준), 인계자 성명, 인수자 성명, 작성일, 인수인계 완료 예정일\n"
        "- tasks: 담당업무 개요 표 (주요 업무별 한 행)\n"
        "- details: 업무별 상세 (tasks 의 각 업무마다 하나). 인수인계 완료 여부는 미완료/진행중/완료 중 하나\n"
        "- urgent_schedule: 긴급·임박한 업무 일정 (일자, 내용, 대응 방법, 담당)\n"
        "- monthly_schedule: 향후 1개월 주요 일정 (일자, 일정 내용, 비고)\n"
        "- communication_note: 대외 커뮤니케이션 유의사항 (여러 줄이면 줄바꿈으로 구분)\n"
        "- assets: 계정·권한·자산 인계 목록\n"
        "- checklist: 인수인계 체크리스트 (자료에 없으면 빈 배열)\n"
        "- signatures: 서명란 이름·날짜 (자료에 명시된 경우만)\n"
        "여러 줄 내용은 ' / ' 또는 줄바꿈으로 구분해 한 칸에 넣는다.\n"
    )
    if mode == "memo":
        return common + (
            "\n입력은 인계자가 작성한 업무메모 한 건이다. 메모의 각 항목을 알맞은 칸에 옮긴다. "
            "메모 형식(키: 값, 표, 자유 문장 등)이 달라도 의미로 판단한다. 메모의 문장은 되도록 원문 그대로 옮긴다.\n"
        )
    return common + (
        "\n입력은 엑셀 업무자료(업무일정·프로젝트·연락망·자산)와 회의록·메일 등 여러 자료다. 이를 종합해 인수인계서를 작성한다.\n"
        "- 인계자·인수자·확인자는 회의록/메일에 명시된 사람을 쓴다. 회사명·부서명은 메일 서명 등에 있으면 쓴다.\n"
        "- tasks/details 는 진행 중인 프로젝트·현장과 반복 업무를 중심으로 구성하고, 각 업무의 담당자·연락처·리스크·참고 파일 경로를 자료에서 찾아 채운다.\n"
        "- urgent_schedule 은 우선순위가 긴급/상이거나 마감이 임박한 항목, monthly_schedule 은 그 외 1개월 내 일정.\n"
        "- assets 는 계정·권한·자산 목록을 그대로 옮긴다(비고에는 완료 목표일과 주의사항).\n"
        "- communication_note 는 연락망의 커뮤니케이션 유의사항을 사람별로 한 줄씩 정리한다.\n"
        "- 작성일·서명일처럼 자료에 없는 값은 비워둔다.\n"
    )


@st.cache_data(show_spinner=False)
def _ai_handover_cached(blob, mode):
    return _call_tool(_handover_system(mode), blob, "fill_handover_form", HANDOVER_SCHEMA, 12000)


def _rows(items, columns):
    rows = []
    for item in items or []:
        if not isinstance(item, dict):
            continue
        row = {col: _s(item.get(col)) for col in columns}
        if any(row.values()):
            rows.append(row)
    return rows


def _merge_handover(output):
    from memo_parser import empty_data

    data = empty_data()
    for group, keys in (("meta", META_KEYS), ("basic", BASIC_KEYS), ("signatures", SIGN_KEYS)):
        source = output.get(group) or {}
        for key in keys:
            data[group][key] = _s(source.get(key))

    data["tasks"] = _rows(output.get("tasks"), TASK_COLUMNS)
    data["details"] = _rows(output.get("details"), DETAIL_FIELDS)
    for detail in data["details"]:
        if detail["인수인계 완료 여부"] not in ("미완료", "진행중", "완료"):
            detail["인수인계 완료 여부"] = "미완료"
    data["urgent_schedule"] = _rows(output.get("urgent_schedule"), URGENT_COLUMNS)
    data["monthly_schedule"] = _rows(output.get("monthly_schedule"), MONTHLY_COLUMNS)
    data["communication_note"] = _s(output.get("communication_note"))
    data["assets"] = _rows(output.get("assets"), ASSET_COLUMNS)

    checklist = _rows(output.get("checklist"), CHECK_COLUMNS)
    if checklist:
        data["checklist"] = checklist  # 없으면 기본 체크리스트 4개 유지

    if not data["tasks"]:
        for detail in data["details"]:
            data["tasks"].append({
                "주요 업무": detail.get("업무명", ""), "업무 목적": detail.get("목적 / 성과지표", ""),
                "업무 프로세스": "", "우선순위": "", "비고": "",
            })

    if not data["signatures"]["인계자 성명"]:
        data["signatures"]["인계자 성명"] = data["basic"]["인계자 성명"]
    if not data["signatures"]["인수자 성명"]:
        data["signatures"]["인수자 성명"] = data["basic"]["인수자 성명"]
    return data


def generate_handover(sources, mode="memo"):
    """
    sources: [(파일명, 텍스트)]
    mode: "memo"(업무메모 1건) / "integrated"(엑셀+문서 종합)
    반환: memo_parser.empty_data() 와 동일한 구조
    """
    blob = "\n\n".join(f"### 자료: {name}\n{text[:15000]}" for name, text in sources)
    output = _ai_handover_cached(blob, mode)
    return _merge_handover(output)
