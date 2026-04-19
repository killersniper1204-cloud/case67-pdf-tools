import os
import re
from datetime import datetime

import pandas as pd
import streamlit as st

# =========================
# 基本設定
# =========================
st.set_page_config(page_title="case66｜T1_2｜df萃取(以前綴增加編號與序號)", layout="wide")
st.title("case66｜T1_2｜df萃取(母版)(以前綴增加編號與序號)")

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SCRIPT_STEM = os.path.splitext(os.path.basename(__file__))[0]
OUT_DIR = os.path.join(SCRIPT_DIR, "out")

# =========================
# 工具函式
# =========================
def safe_filename(name: str) -> str:
    name = str(name).strip()
    if not name:
        name = "output"
    return re.sub(r'[\\/*?:"<>|]+', "_", name)


def safe_sheet_name(name: str, used_names=None) -> str:
    if used_names is None:
        used_names = set()

    s = str(name).strip()
    if not s:
        s = "Sheet"

    s = re.sub(r'[\\/*?:\[\]]+', "_", s)
    s = s[:31] if len(s) > 31 else s
    if not s:
        s = "Sheet"

    base = s
    i = 1
    while s in used_names:
        suffix = f"_{i}"
        s = (base[:31 - len(suffix)] + suffix) if len(base) + len(suffix) > 31 else base + suffix
        i += 1

    used_names.add(s)
    return s


def clean_prefix(s: str) -> str:
    if s is None:
        return ""
    return str(s).strip()


def clean_text(x) -> str:
    if pd.isna(x):
        return ""
    return str(x).strip()


def strip_number_dot_prefix(text: str) -> str:
    """
    去除前綴如：
    1. xxx
    2. xxx
    10. xxx
    """
    text = clean_text(text)
    return re.sub(r"^\s*\d+\.\s*", "", text)


# =========================
# 核心演算 1：以前綴增加編號與序號
# =========================
def build_numbering_by_prefix(df: pd.DataFrame, target_col: str, prefix: str) -> pd.DataFrame:
    """
    規則：
    1. 第一筆固定：編號=1、序號=1
    2. 從第二筆開始往下演算
    3. 若該列文字以前綴 prefix 開頭，則視為新群組起點，編號 += 1，序號 = 1
    4. 若未命中前綴，則延續同一編號，序號累加
    """
    df = df.copy()

    n = len(df)
    if n == 0:
        df["編號"] = []
        df["序號"] = []
        return df

    group_ids = []
    seq_ids = []

    current_group = 1
    current_seq = 1
    group_ids.append(current_group)
    seq_ids.append(current_seq)

    for i in range(1, n):
        val = df.iloc[i][target_col] if target_col in df.columns else ""
        text = clean_text(val)

        if prefix != "" and text.startswith(prefix):
            current_group += 1
            current_seq = 1
        else:
            current_seq += 1

        group_ids.append(current_group)
        seq_ids.append(current_seq)

    df["編號"] = group_ids
    df["序號"] = seq_ids
    return df


# =========================
# 核心演算 2：以編號範圍回填 type
# =========================
def fill_type_by_group(df: pd.DataFrame, target_col: str) -> pd.DataFrame:
    """
    規則：
    1. 以編號為範圍巡覽
    2. 當找到內容以 '1.' 開頭的列
    3. 取該列上一列內容作為 type 值
    4. 將該 type 填入此編號範圍內所有列的 type 欄位
    """
    df = df.copy()
    df["type"] = ""

    if "編號" not in df.columns or target_col not in df.columns or df.empty:
        return df

    for _, idxs in df.groupby("編號", sort=False).groups.items():
        idx_list = list(idxs)
        type_value = ""

        for pos, idx in enumerate(idx_list):
            text = clean_text(df.at[idx, target_col])
            if text.startswith("1."):
                if pos - 1 >= 0:
                    prev_idx = idx_list[pos - 1]
                    type_value = clean_text(df.at[prev_idx, target_col])
                break

        if type_value != "":
            df.loc[idx_list, "type"] = type_value

    return df


# =========================
# 核心演算 3：以編號範圍產生 新文本
# =========================
def fill_new_text_by_group(df: pd.DataFrame, target_col: str) -> pd.DataFrame:
    """
    規則：
    1. 以編號為範圍巡覽
    2. 當找到內容以 '1.' 開頭的列，視為新文本起點
    3. 從該列到範圍末列，去除數字點前綴後串接成新文本
    4. 新文本只填入該編號中『序號=1』那一列
    """
    df = df.copy()
    df["新文本"] = ""

    if "編號" not in df.columns or "序號" not in df.columns or target_col not in df.columns or df.empty:
        return df

    for _, idxs in df.groupby("編號", sort=False).groups.items():
        idx_list = list(idxs)
        start_pos = None

        for pos, idx in enumerate(idx_list):
            text = clean_text(df.at[idx, target_col])
            if text.startswith("1."):
                start_pos = pos
                break

        if start_pos is None:
            continue

        parts = []
        for idx in idx_list[start_pos:]:
            text = clean_text(df.at[idx, target_col])
            text = strip_number_dot_prefix(text)
            if text != "":
                parts.append(text)
        # ⭐ 補丁：編號=1 時，加上「，」
        if len(parts) > 0:
            if df.at[idx_list[0], "編號"] == 1:
                new_text = "，".join(parts).strip()
            else:
                new_text = "".join(parts).strip()
        else:
            new_text = ""
        if new_text == "":
            continue

        first_seq_idx = None
        for idx in idx_list:
            seq_val = df.at[idx, "序號"]
            try:
                if int(seq_val) == 1:
                    first_seq_idx = idx
                    break
            except Exception:
                pass

        if first_seq_idx is not None:
            df.at[first_seq_idx, "新文本"] = new_text

    return df


# =========================
# 核心演算 4：tx 解析成 dfx
# =========================
def parse_kv_pairs(text: str):
    """
    從字串中抓出連續的 k：v pairs
    支援：
    - K1：v1，K2：v2，K3：v3
    - K1：v1,K2：v2,K3：v3
    - K1：v1；K2：v2；K3：v3
    - K1：v1、K2：v2、K3：v3
    - 甲：01 乙：02甲：11 乙：12

    回傳：
    [('K1','v1'), ('K2','v2'), ...]
    """
    text = clean_text(text)
    if text == "":
        return []

    s = text.replace(":", "：")

    # 先在明顯分隔符後補換行
    s = s.replace("，", "\n")
    s = s.replace(",", "\n")
    s = s.replace("；", "\n")
    s = s.replace(";", "\n")
    s = s.replace("、", "\n")

    # 把「空白 + 新key：」補切開，支援：
    # 甲：01 乙：02
    s = re.sub(r'\s+(?=[^：\s]+：)', "\n", s)

    lines = [clean_text(x) for x in s.split("\n") if clean_text(x) != ""]

    pairs = []

    for line in lines:
        # 逐段抓 line 中所有 k：v
        found = re.findall(r'([^：\s]+)：(.*?)(?=(?:[^：\s]+)：|$)', line)
        for k, v in found:
            k = clean_text(k)
            v = clean_text(v)
            if k != "":
                pairs.append((k, v))

    return pairs


def split_tx_by_first_key(tx: str):
    """
    以 tx 中第一個出現的 key 當作檢索值，
    以該 key： 作為每筆資料的起點切成多個 tx2。

    例如：
    甲：01 乙：02甲：11 乙：12
    =>
    ['甲：01 乙：02', '甲：11 乙：12']

    也支援：
    甲：01，乙：02，甲：11，乙：12
    """
    tx = clean_text(tx)
    if tx == "":
        return []

    pairs = parse_kv_pairs(tx)
    if not pairs:
        return [tx]

    first_key = pairs[0][0]
    marker = f"{first_key}："

    # 先統一常見分隔符，幫助起點切割更穩
    s = tx.replace(":", "：")
    s = s.replace("，", " ")
    s = s.replace(",", " ")
    s = s.replace("；", " ")
    s = s.replace(";", " ")
    s = s.replace("、", " ")

    starts = [m.start() for m in re.finditer(re.escape(marker), s)]
    if not starts:
        return [tx]

    tx2_list = []
    for i, start in enumerate(starts):
        end = starts[i + 1] if i + 1 < len(starts) else len(s)
        seg = clean_text(s[start:end])
        if seg != "":
            tx2_list.append(seg)

    return tx2_list if tx2_list else [tx]


def tx_to_dfx(tx: str) -> pd.DataFrame:
    """
    將 tx 解析成 dfx

    規則：
    1. 先抓出第一個 key，作為檢索值
    2. 以「檢索值：」切成多筆 tx2
    3. 每筆 tx2 再解析成多個 k:v
    4. 動態生成欄位，形成 dfx

    範例：
    甲：01 乙：02甲：11 乙：12
    =>
    id | tx2            | 甲 | 乙
    1  | 甲：01 乙：02   | 01 | 02
    2  | 甲：11 乙：12   | 11 | 12

    K1：v1，K2：v2，K1：v3，K2：v4
    =>
    id | tx2                  | K1 | K2
    1  | K1：v1 K2：v2        | v1 | v2
    2  | K1：v3 K2：v4        | v3 | v4
    """
    tx2_list = split_tx_by_first_key(tx)

    rows = []
    dynamic_cols = []

    for i, tx2 in enumerate(tx2_list, start=1):
        row = {"id": i, "tx2": tx2}

        pairs = parse_kv_pairs(tx2)
        for k, v in pairs:
            row[k] = v
            if k not in dynamic_cols:
                dynamic_cols.append(k)

        rows.append(row)

    dfx = pd.DataFrame(rows)

    final_cols = ["id", "tx2"] + dynamic_cols
    for c in final_cols:
        if c not in dfx.columns:
            dfx[c] = ""

    dfx = dfx[final_cols]
    return dfx


def build_dfx_sheets_from_result_df(result_df: pd.DataFrame):
    """
    由上巡覽 新文本 欄位。
    當新文本有值時：
    - 取該筆 type 作為工作表名
    - 取該筆 新文本 作為 tx
    - 解析成 dfx
    """
    sheets = {}
    used_sheet_names = set()

    if result_df is None or result_df.empty:
        return sheets

    for _, row in result_df.iterrows():
        tx = clean_text(row.get("新文本", ""))
        if tx == "":
            continue

        type_name = clean_text(row.get("type", ""))
        if type_name == "":
            type_name = "Sheet"

        dfx = tx_to_dfx(tx)
        sheet_name = safe_sheet_name(type_name, used_sheet_names)
        sheets[sheet_name] = dfx

    return sheets


# =========================
# Session state
# =========================
if "t1_2_df" not in st.session_state:
    st.session_state.t1_2_df = None

if "t1_2_result_df" not in st.session_state:
    st.session_state.t1_2_result_df = None

if "t1_2_dfx_sheets" not in st.session_state:
    st.session_state.t1_2_dfx_sheets = {}

if "t1_2_sheet_name" not in st.session_state:
    st.session_state.t1_2_sheet_name = None

if "t1_2_file_name" not in st.session_state:
    st.session_state.t1_2_file_name = None

if "t1_2_run_ok" not in st.session_state:
    st.session_state.t1_2_run_ok = False


# =========================
# Sidebar
# =========================
st.sidebar.header("工序設定")

uploaded_file = st.sidebar.file_uploader("1. 上傳 .xlsx", type=["xlsx"])

sheet_names = []
xls = None

if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
    except Exception as e:
        st.sidebar.error(f"讀取 Excel 失敗：{e}")

selected_sheet = st.sidebar.selectbox(
    "2. 下拉選定工作表當 df",
    options=sheet_names if sheet_names else [""],
    index=0
)

df_preview = None
if uploaded_file is not None and selected_sheet:
    try:
        df_preview = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
    except Exception as e:
        st.error(f"讀取工作表失敗：{e}")

columns = list(df_preview.columns) if df_preview is not None else []
target_col = st.sidebar.selectbox(
    "3. 選定內容分析欄位",
    options=columns if columns else [""],
    index=0
)

prefix_value = st.sidebar.text_input("4. 前綴特徵值", value="")

run_btn = st.sidebar.button("執行")
export_btn = st.sidebar.button("匯出")


# =========================
# 執行
# =========================
if run_btn:
    if uploaded_file is None:
        st.error("請先上傳 .xlsx")
    elif not selected_sheet:
        st.error("請先選定工作表")
    elif df_preview is None:
        st.error("目前 df 讀取失敗")
    elif not target_col:
        st.error("請先選定內容分析欄位")
    else:
        prefix = clean_prefix(prefix_value)

        result_df = build_numbering_by_prefix(df_preview, target_col, prefix)
        result_df = fill_type_by_group(result_df, target_col)
        result_df = fill_new_text_by_group(result_df, target_col)
        dfx_sheets = build_dfx_sheets_from_result_df(result_df)

        st.session_state.t1_2_df = df_preview.copy()
        st.session_state.t1_2_result_df = result_df
        st.session_state.t1_2_dfx_sheets = dfx_sheets
        st.session_state.t1_2_sheet_name = selected_sheet
        st.session_state.t1_2_file_name = uploaded_file.name
        st.session_state.t1_2_run_ok = True


# =========================
# 主畫面顯示
# =========================
st.markdown("### 說明")
st.write(
    "第一筆固定為編號1、序號1；從第二筆開始依前綴特徵值逐段增加『編號』與『序號』，"
    "再以編號範圍內找到 '1.' 的列，將其上一列內容填入同範圍的 type 欄位；"
    "並從 '1.' 起到範圍末列，去除數字點前綴後串接為『新文本』，只填入該編號的序號1列；"
    "最後由上巡覽『新文本』，以該列 type 作為表名，並以第一個 key 作為檢索值切成多筆 tx2，再解析成 dfx。"
)

if df_preview is not None:
    st.markdown("### 原始 df")
    st.dataframe(df_preview, use_container_width=True)

if st.session_state.t1_2_run_ok and st.session_state.t1_2_result_df is not None:
    result_df = st.session_state.t1_2_result_df
    st.success(f"處理完成。result_df.shape = {result_df.shape}")

    st.markdown("### 結果 df")
    st.dataframe(result_df, use_container_width=True)

    dfx_sheets = st.session_state.t1_2_dfx_sheets
    if dfx_sheets:
        st.markdown("### dfx 預覽")
        for sheet_name, dfx in dfx_sheets.items():
            st.markdown(f"#### {sheet_name}")
            st.dataframe(dfx, use_container_width=True)


# =========================
# 匯出
# =========================
if export_btn:
    if not st.session_state.t1_2_run_ok or st.session_state.t1_2_result_df is None:
        st.error("請先按執行，產生結果後再匯出。")
    else:
        try:
            os.makedirs(OUT_DIR, exist_ok=True)

            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            src_name = os.path.splitext(st.session_state.t1_2_file_name or "input")[0]
            src_name = safe_filename(src_name)

            out_name = f"{SCRIPT_STEM}_{src_name}_{ts}_result.xlsx"
            out_path = os.path.join(OUT_DIR, out_name)

            with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
                st.session_state.t1_2_result_df.to_excel(
                    writer,
                    sheet_name="result",
                    index=False
                )

                for sheet_name, dfx in st.session_state.t1_2_dfx_sheets.items():
                    dfx.to_excel(writer, sheet_name=sheet_name, index=False)

            st.success(f"匯出成功：{out_path}")
        except Exception as e:
            st.error(f"匯出失敗：{e}")

if st.session_state.t1_2_result_df is None:
    st.info("請依序設定參數，按『執行』產生結果；按『匯出』才寫入 \\out。")