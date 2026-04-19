import streamlit as st
import pandas as pd
from pypdf import PdfReader
import os
import re
import time
from datetime import datetime
import tempfile

# ======================
# 基本設定
# ======================
st.set_page_config(page_title="case66｜PDF逐列文字轉df", layout="wide")
st.title("case66｜T1_1｜PDF逐列文字 → df(page, text)")

# ======================
# 常數 / regex
# ======================
CONTROL_RE = re.compile(r"[\x00-\x08\x0b-\x0c\x0e-\x1f]")

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(SCRIPT_DIR, "out")
SCRIPT_STEM = os.path.splitext(os.path.basename(__file__))[0]

# ======================
# 工具函式
# ======================
def excel_safe_text(s):
    """移除 Excel XML 不允許字元"""
    if not isinstance(s, str):
        return s

    cleaned = []
    for ch in s:
        code = ord(ch)
        if (
            code == 0x9 or
            code == 0xA or
            code == 0xD or
            0x20 <= code <= 0xD7FF or
            0xE000 <= code <= 0xFFFD or
            0x10000 <= code <= 0x10FFFF
        ):
            cleaned.append(ch)
    return "".join(cleaned)

def parse_pages(pages_str, total_pages):
    """解析頁碼字串，例如 all、1,3,5-8"""
    if not pages_str or pages_str.strip().lower() == "all":
        return list(range(1, total_pages + 1))

    pages = set()
    parts = pages_str.split(",")

    for p in parts:
        p = p.strip()
        if not p:
            continue

        if "-" in p:
            ab = p.split("-", 1)
            if len(ab) != 2:
                continue
            try:
                a = int(ab[0].strip())
                b = int(ab[1].strip())
            except ValueError:
                continue

            if a > b:
                a, b = b, a

            for x in range(a, b + 1):
                if 1 <= x <= total_pages:
                    pages.add(x)
        else:
            try:
                x = int(p)
            except ValueError:
                continue
            if 1 <= x <= total_pages:
                pages.add(x)

    return sorted(pages)

def clean_df_for_excel(df):
    if df is None or df.empty:
        return df

    df = df.copy()
    df.columns = [excel_safe_text(CONTROL_RE.sub("", str(c))) for c in df.columns]

    for c in df.columns:
        df[c] = df[c].apply(
            lambda x: excel_safe_text(CONTROL_RE.sub("", x)) if isinstance(x, str) else x
        )
    return df

def pdf_to_text_rows(pdf_path, pages_str):
    start = time.time()
    logs = []

    try:
        reader = PdfReader(pdf_path)
    except Exception as e:
        return None, [f"❌ 無法讀取 PDF：{e}"], 0.0

    total_pages = len(reader.pages)
    logs.append(f"PDF總頁數：{total_pages}")

    pages = parse_pages(pages_str, total_pages)
    logs.append(f"解析頁碼：{pages}")

    rows = []
    for p in pages:
        try:
            text = reader.pages[p - 1].extract_text()
        except Exception as e:
            logs.append(f"⚠ 第 {p} 頁讀取失敗：{e}")
            continue

        if not text:
            continue

        for line in text.split("\n"):
            line = CONTROL_RE.sub("", line.strip())
            line = excel_safe_text(line)
            if line:
                rows.append([p, line])

    df = pd.DataFrame(rows, columns=["page", "text"])
    df = clean_df_for_excel(df)

    elapsed = time.time() - start
    logs.append(f"df.shape = {df.shape}")
    logs.append(f"完成解析，用時 {elapsed:.2f} 秒")

    return df, logs, elapsed

def export_to_xlsx(df, out_name):
    os.makedirs(OUT_DIR, exist_ok=True)
    out_path = os.path.join(OUT_DIR, out_name)

    df_export = clean_df_for_excel(df.copy())
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df_export.to_excel(writer, index=False, sheet_name="text_rows")

    return out_path

# ======================
# Session State
# ======================
if "result_df" not in st.session_state:
    st.session_state.result_df = None

if "result_logs" not in st.session_state:
    st.session_state.result_logs = []

if "result_elapsed" not in st.session_state:
    st.session_state.result_elapsed = 0.0

if "result_file_name" not in st.session_state:
    st.session_state.result_file_name = ""

# ======================
# 側邊欄
# ======================
st.sidebar.header("參數設定")

uploaded_file = st.sidebar.file_uploader("上傳 PDF", type=["pdf"])
pages_str = st.sidebar.text_input("頁碼（預設 all）", value="all")

default_name = f"{SCRIPT_STEM}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
output_name = st