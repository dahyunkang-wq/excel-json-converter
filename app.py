%%writefile app.py
import io
import json
import re
import zipfile

import pandas as pd
import streamlit as st


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


st.title("엑셀 (D12~F열) → JSON txt 변환기 (Colab + Streamlit)")
st.write("같은 포맷의 엑셀 파일 여러 개를 업로드하면, 각 파일을 JSON으로 변환해서 다운로드할 수 있습니다.")

uploaded_files = st.file_uploader(
    "엑셀 파일(.xlsx, .xls)을 하나 이상 선택하세요",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
)

if uploaded_files:
    all_json_strings = {}
    st.subheader("변환 결과 미리보기")

    for file in uploaded_files:
        st.markdown(f"### 파일: **{file.name}**")

        try:
            df = pd.read_excel(file, header=None)
        except Exception as e:
            st.error(f"{file.name} 읽기 실패: {e}")
            continue

        records = excel_to_json_records(df)
        json_str = json.dumps(records, ensure_ascii=False, indent=2)

        all_json_strings[file.name] = json_str

        st.code(json_str, language="json")

        base_name = file.name.rsplit(".", 1)[0]
        st.download_button(
            label=f"{file.name} → JSON txt 다운로드",
            data=json_str.encode("utf-8"),
            file_name=f"{base_name}.json.txt",
            mime="text/plain",
        )

    if len(all_json_strings) > 1:
        st.subheader("ZIP으로 한 번에 받기")

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for fname, jstr in all_json_strings.items():
                base_name = fname.rsplit(".", 1)[0]
                zf.writestr(f"{base_name}.json.txt", jstr)

        zip_buffer.seek(0)
        st.download_button(
            label="모든 JSON txt 파일 ZIP 다운로드",
            data=zip_buffer,
            file_name="json_outputs.zip",
            mime="application/zip",
        )
else:
    st.info("오른쪽 파일 업로더를 통해 엑셀 파일을 올려주세요.")
