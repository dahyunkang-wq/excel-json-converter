# -*- coding: utf-8 -*-
import io
import json
import re
import zipfile
import base64
from io import BytesIO
from pathlib import Path
from typing import List, Dict, Any, Tuple
import unicodedata  # 한글 자모 조합(NFC)을 위해 추가

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

# openpyxl 및 스타일 관련 모듈 추가
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.styles.borders import Border, Side
from openpyxl.styles.colors import Color
from openpyxl.text.rich_text import RichText
from openpyxl.cell.text import Text


# =============================================================================
#
# 스크립트 1 (Excel → JSON) 헬퍼 함수
# (이전과 동일)
#
# =============================================================================

def normalize_category_name(raw_key: str) -> str:
    key = raw_key.strip().lower()
    key = key.replace(":", "")
    key = re.sub(r"\s+", "_", key)

    mapping = {
        "language": "language",
        "languages": "language",
        "audio_processing": "audio_processing",
        "audio": "audio_processing",
        "data_handling": "data_handling",
        "data": "data_handling",
        "tools": "tools",
        "tool": "tools",
    }

    return mapping.get(key, "etc")


def split_items(text: str):
    if not isinstance(text, str):
        return []

    parts = re.split(r"[\n,\r,]+", text)
    cleaned = []
    for p in parts:
        p = p.strip()
        p = re.sub(r'^[\*\-\·\u2022]+\s*', "", p)
        if p:
            cleaned.append(p)
    return cleaned


def parse_tech_stack(raw_text: str):
    result = {
        "language": [],
        "audio_processing": [],
        "data_handling": [],
        "tools": [],
        "etc": [],
    }

    if not isinstance(raw_text, str) or not raw_text.strip():
        return result

    lines = re.split(r"[\r\n]+", raw_text)
    current_key = None
    etc_buffer = []

    for line in lines:
        if not line or not line.strip():
            continue

        line = re.sub(r'^[\*\-\·\u2022]+\s*', "", line).strip()
        if not line:
            continue

        if ":" in line:
            raw_key, value = line.split(":", 1)
            cat = normalize_category_name(raw_key)
            current_key = cat

            items = split_items(value)
            if cat in result:
                result[cat].extend(items)
            else:
                result["etc"].extend(items)
        else:
            items = split_items(line)
            if current_key and current_key in result:
                result[current_key].extend(items)
            elif current_key and current_key not in result:
                result["etc"].extend(items)
            else:
                etc_buffer.extend(items)

    if etc_buffer:
        result["etc"].extend(etc_buffer)

    for key in list(result.keys()):
        seen = set()
        unique_items = []
        for item in result[key]:
            item = item.strip()
            if not item:
                continue
            if item in seen:
                continue
            seen.add(item)
            unique_items.append(item)
        result[key] = unique_items

    return result


def clean_task_description(raw_text: str) -> str:
    if not isinstance(raw_text, str):
        raw_text = str(raw_text) if raw_text is not None else ""
    text = re.sub(r"\s+", " ", raw_text).strip()
    return text


def excel_to_json_records(df: pd.DataFrame):
    records = []

    start_row = 11  # 12행
    num_rows = df.shape[0]

    for i in range(start_row, num_rows):
        d_val = df.iloc[i, 3] if df.shape[1] > 3 else None
        e_val = df.iloc[i, 4] if df.shape[1] > 4 else None
        f_val = df.iloc[i, 5] if df.shape[1] > 5 else None

        def is_empty(v):
            if v is None:
                return True
            if isinstance(v, float) and pd.isna(v):
                return True
            if isinstance(v, str) and not v.strip():
                return True
            return False

        if is_empty(d_val) and is_empty(e_val) and is_empty(f_val):
            break

        task_name = "" if d_val is None else str(d_val).strip()
        task_description = clean_task_description(e_val)
        tech_stack = parse_tech_stack("" if f_val is None else str(f_val))

        records.append(
            {
                "task_name": task_name,
                "task_description": task_description,
                "tech_stack": tech_stack,
            }
        )

    return records


# =============================================================================
#
# 스크립트 2 (JSON → Excel) 헬퍼 함수
#
# =============================================================================

# ==========================
# 상수 / 경로
# ==========================
# Streamlit에서 __file__은 스크립트 경로를 올바르게 참조합니다.
APP_DIR = Path(__file__).parent
TEMPLATE_DIR = APP_DIR / "templates"
DEFAULT_TEMPLATE_NONTRACK = "Non Track_Paper Interview_상위조직명_직무명(포맷).xlsx"
DEFAULT_TEMPLATE_TRACK    = "Track_Paper Interview_상위조직명_직무명(포맷).xlsx"

# Non Track 쓰기 범위
TASK_START_ROW_NT, TASK_END_ROW_NT   = 5, 14    # Task: A(이름), C(설명)
SKILL_START_ROW_NT, SKILL_END_ROW_NT = 5, 11    # Skill: A/B/D/F

# Track 쓰기 범위 (규칙 동일)
TASK_ROW_START_T, TASK_ROW_END_T   = 5, 14
SKILL_ROW_START_T, SKILL_ROW_END_T = 5, 11
TASK_TEMPLATE_SHEET_T  = "Task"
SKILL_TEMPLATE_SHEET_T = "Skill"
TRACK_TITLE_RANGE_T    = "D1:D2"  # 트랙명 표기 영역

# ==========================
# 공통: 텍스트 정리(마커 제거)
# ==========================
# [cite: ...]
CITE_PATTERN = re.compile(r'\s*\[\s*cite\s*:\s*.*?\]\s*', flags=re.IGNORECASE | re.DOTALL)
# (Source ...)
SOURCE_PAREN_PATTERN = re.compile(r'\s*\(\s*source[^)]*\)\s*', flags=re.IGNORECASE)

def strip_markers(text: Any) -> str:
    """[cite: ...], (Source ...) 제거 + 공백 정리"""
    if text is None:
        return ""
    s = str(text)
    s = CITE_PATTERN.sub(" ", s)
    s = SOURCE_PAREN_PATTERN.sub(" ", s)
    s = re.sub(r"[ \t]+", " ", s).strip()
    return s

# ==========================
# 공통: 파일명 유틸
# ==========================
INVALID_WIN_CHARS = r'<>:"/\\|?*'
INVALID_WIN_PATTERN = re.compile(f"[{re.escape(INVALID_WIN_CHARS)}]+")

def sanitize_filename_component(s: str, fallback: str = "untitled") -> str:
    if not s:
        return fallback
    s = INVALID_WIN_PATTERN.sub(" ", s).strip().strip(".")
    return s if s else fallback

# ==========================
# Non Track 파서/로직
# ==========================
def title_tokens_nt(stem: str) -> List[str]:
    return [t.strip() for t in stem.split("_") if t.strip()]

def is_trailing_excluded_nt(token: str) -> bool:
    t = token.lower().replace(" ", "")
    return t in {"skill", "hc제외"}

def parse_org_role_from_filename_nt(filename: str) -> Tuple[str, str, str]:
    """{상위조직명} = 첫 토큰, {직무명} = 두 번째~끝(뒤에서 skill/HC 제외 제거), 표시/파일명 둘 다 '공백' 연결"""
    stem = Path(filename).stem
    toks = title_tokens_nt(stem)
    if not toks:
        return "unknown", "", ""
    org = toks[0]
    end = len(toks)
    while end > 1 and is_trailing_excluded_nt(toks[end - 1]):
        end -= 1
    role_tokens = toks[1:end] if end > 1 else toks[1:]
    role_display = " ".join(role_tokens)
    role_for_filename = " ".join(role_tokens)
    return org, role_display, role_for_filename

def with_wrap(cell):
    a = cell.alignment or Alignment()
    return Alignment(
        horizontal=a.horizontal,
        vertical=a.vertical,
        text_rotation=a.text_rotation,
        wrap_text=True,
        shrink_to_fit=a.shrink_to_fit,
        indent=a.indent
    )

def set_text(ws, coord: str, text: str, wrap: bool = True):
    cell = ws[coord]
    cell.value = text
    if wrap:
        cell.alignment = with_wrap(cell)

def load_json_from_txt_bytes(b: bytes) -> Dict[str, Any]:
    """TXT에 전후 텍스트가 섞여 있어도 {} 블록만 추출 시도"""
    txt = b.decode("utf-8-sig", errors="ignore")
    try:
        return json.loads(txt)
    except json.JSONDecodeError:
        start = txt.find("{")
        end = txt.rfind("}")
        if start != -1 and end != -1 and start < end:
            return json.loads(txt[start:end+1])
        raise

def collect_tasks_nt(obj: Dict[str, Any]) -> List[Dict[str, Any]]:
    return obj.get("tasks") or []

def iter_skills_nt(obj: Dict[str, Any]):
    skills = obj.get("skills") or []
    for item in skills:
        if isinstance(item, dict) and "skill" in item:
            s = item.get("skill") or {}
            name = s.get("name", "")
            definition = s.get("definition", "")
            stack = s.get("tech_stack", {})
            related = item.get("related_tasks") or s.get("related_tasks") or []
        else:
            s = item if isinstance(item, dict) else {}
            name = s.get("name", "")
            definition = s.get("definition", "")
            stack = s.get("tech_stack", {})
            related = s.get("related_tasks") or []
        yield {"name": name, "definition": definition, "tech_stack": stack, "related_tasks": related}

def normalize_list(val) -> List[str]:
    if val is None:
        return []
    if isinstance(val, (list, tuple, set)):
        return [str(x).strip() for x in val if str(x).strip()]
    s = str(val).strip()
    if not s:
        return []
    parts = []
    for chunk in s.replace(";", ",").replace("/", ",").split(","):
        chunk = chunk.strip()
        if chunk:
            parts.append(chunk)
    return parts

def extract_tech_lines_nt(tech_stack: Dict[str, Any]) -> str:
    if not isinstance(tech_stack, dict):
        tech_stack = {}
    lower_map = {str(k).lower(): v for k, v in tech_stack.items()}
    languages = normalize_list(lower_map.get("language") or lower_map.get("languages"))
    os_list   = normalize_list(lower_map.get("os") or lower_map.get("platform") or lower_map.get("operating_system"))
    tools     = normalize_list(lower_map.get("tools") or lower_map.get("tool"))
    lines = []
    if languages: lines.append(f"* language: {', '.join(languages)}")
    if os_list:   lines.append(f"* os: {', '.join(os_list)}")
    if tools:     lines.append(f"* tools: {', '.join(tools)}")
    return strip_markers("\n".join(lines))  # ← 마커 제거

def bullet_lines(items: List[str]) -> str:
    items = [str(i).strip() for i in items if str(i).strip()]
    return "\n".join(f"* {i}" for i in items)

def related_task_names_nt(related_tasks: List[Dict[str, Any]], task_id_to_name: Dict[str, str]) -> List[str]:
    names = []
    for rt in related_tasks or []:
        name = (rt.get("task_name") or "").strip()
        if not name:
            tid = (rt.get("task_id") or "").strip()
            if tid and tid in task_id_to_name:
                name = task_id_to_name[tid]
        if name:
            names.append(name)
    return names

def build_workbook_nontrack(template_bytes: bytes, org: str, role: str, data: Dict[str, Any]) -> BytesIO:
    """템플릿 서식 유지, 값만 주입"""
    wb = load_workbook(BytesIO(template_bytes))
    ws_task  = wb["Task"] if "Task" in wb.sheetnames else wb[wb.sheetnames[0]]
    ws_skill = wb["Skill"] if "Skill" in wb.sheetnames else wb[wb.sheetnames[1]]

    # Task
    set_text(ws_task, "B1", org) # B1, B2는 VBA 수정 함수에서 한글 교정됨
    set_text(ws_task, "B2", role)
    tasks = collect_tasks_nt(data)
    task_id_to_name = {}
    for t in tasks:
        tid = str(t.get("task_id") or "").strip()
        tname = str(t.get("task_name") or "").strip()
        if tid and tname:
            task_id_to_name[tid] = tname
    row = TASK_START_ROW_NT
    for t in tasks[: (TASK_END_ROW_NT - TASK_START_ROW_NT + 1) ]:
        set_text(ws_task, f"A{row}", str(t.get("task_name") or "").strip())
        set_text(ws_task, f"C{row}", str(t.get("task_description") or "").strip())
        row += 1
    for r in range(row, TASK_END_ROW_NT + 1):
        set_text(ws_task, f"A{r}", ""); set_text(ws_task, f"C{r}", "")

    # Skill
    set_text(ws_skill, "B1", org) # B1, B2는 VBA 수정 함수에서 한글 교정됨
    set_text(ws_skill, "B2", role)
    processed = 0
    max_rows = SKILL_END_ROW_NT - SKILL_START_ROW_NT + 1
    for s in iter_skills_nt(data):
        if processed >= max_rows: break
        r = SKILL_START_ROW_NT + processed
        rel_names = related_task_names_nt(s.get("related_tasks"), task_id_to_name)
        set_text(ws_skill, f"A{r}", bullet_lines(rel_names) if rel_names else "")
        set_text(ws_skill, f"B{r}", str(s.get("name") or "").strip())
        set_text(ws_skill, f"D{r}", strip_markers(s.get("definition")))
        set_text(ws_skill, f"F{r}", extract_tech_lines_nt(s.get("tech_stack")))
        processed += 1
    for r in range(SKILL_START_ROW_NT + processed, SKILL_END_ROW_NT + 1):
        for c in ("A","B","D","F"):
            set_text(ws_skill, f"{c}{r}", "")

    # --- VBA 스타일 적용 ---
    apply_vba_description_edits(wb)
    apply_vba_extra_borders_and_dims(wb)
    apply_vba_global_font(wb, "현대하모니 L")
    apply_vba_korean_fix_to_headers(wb) # B1, B2 한글 교정
    # --- ---

    bio = BytesIO(); wb.save(bio); bio.seek(0); return bio

def process_uploaded_txt_nontrack(uploaded_file, template_bytes: bytes):
    org, role_display, role_for_filename = parse_org_role_from_filename_nt(uploaded_file.name)
    safe_org  = sanitize_filename_component(org, "org")
    safe_role = sanitize_filename_component(role_for_filename, "role")
    out_name = f"Non Track_Paper Interview_{safe_org}_{safe_role}.xlsx"
    data = load_json_from_txt_bytes(uploaded_file.read())
    # build_workbook_nontrack 내부에서 VBA 스타일 적용
    wb_bytes = build_workbook_nontrack(template_bytes, org, role_display, data)
    return out_name, wb_bytes

# ==========================
# Track 파서/로직
# ==========================
def parse_org_and_job_from_filename_track(filename: str) -> Tuple[str, str]:
