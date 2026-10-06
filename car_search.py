import streamlit as st
import pandas as pd
import openpyxl
import re
import os
from datetime import datetime
from PIL import Image

# ページの設定
st.set_page_config(
    page_title="車両管理＆検索アプリ", 
    page_icon="🚗", 
    layout="wide"
)

# ── スマホでも崩れない3列キープCSS ──
st.markdown("""
    <style>
    h1 {
        font-size: 1.45rem !important;
        word-break: break-all;
    }
    [data-testid="stDataFrame"] {
        width: 100% !important;
    }
    div[data-testid="stHorizontalBlock"] {
        display: flex !important;
        flex-direction: row !important;
        flex-wrap: nowrap !important;
        gap: 6px !important;
    }
    div[data-testid="stHorizontalBlock"] > div {
        flex: 1 1 0% !important;
        min-width: 0 !important;
    }
    div[data-testid="stHorizontalBlock"] button, div[data-testid="stHorizontalBlock"] a {
        width: 100% !important;
        height: 48px !important;
        font-size: 1.05rem !important;
        font-weight: bold !important;
        border-radius: 8px !important;
        padding: 0 !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
    }
    .vehicle-detail-card {
        background-color: #f8f9fa;
        border: 2px solid #1e88e5;
        border-radius: 12px;
        padding: 16px;
        margin-top: 15px;
        margin-bottom: 20px;
        box-shadow: 0 4px 10px rgba(0,0,0,0.08);
    }
    .card-title {
        font-size: 1.35rem;
        font-weight: bold;
        color: #1e88e5;
        margin-bottom: 8px;
        border-bottom: 2px solid #e0e0e0;
        padding-bottom: 4px;
    }
    .card-row {
        font-size: 1.05rem;
        margin: 6px 0;
        color: #222;
    }
    .card-label {
        font-weight: bold;
        color: #555;
        display: inline-block;
        width: 85px;
    }
    </style>
""", unsafe_allow_html=True)

FILE_PATH = "新_登録車両資料_連動版.xlsx"
SHEET_NAME = "新_登録車両資料"

# 安全な文字列変換関数（floatやNaN、Noneを完全排除）
def safe_str(val):
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    if s.lower() in ["none", "nan", "<na>"]:
        return ""
    return s

# 文字列の表記ゆれ統一
def normalize_text(text):
    s = safe_str(text)
    if not s:
        return ""
    t = s.replace("：", ":").replace(" ", "").replace(" ", "")
    return t.strip()

# データの初期ロード処理
def load_raw_data():
    if not os.path.exists(FILE_PATH):
        return pd.DataFrame(columns=['会社名', '車番', '入力番号', '詳細', '備考', '風体'])

    try:
        raw_df = pd.read_excel(FILE_PATH, sheet_name=SHEET_NAME, header=None)
    except Exception:
        raw_df = pd.read_excel(FILE_PATH, header=None)

    raw_df = raw_df.dropna(how="all").astype(str)
    raw_df = raw_df.replace(r'^\s*$', pd.NA, regex=True)

    processed_rows = []
    current_comp = ""
    
    for _, row in raw_df.iterrows():
        row_vals = [safe_str(val) for val in row.values if safe_str(val)]
        row_text = " ".join(row_vals)
        
        if ":" in row_text or "：" in row_text:
            for val in row_vals:
                if ":" in val or "：" in val:
                    current_comp = val.strip()
                    break
            continue 
        
        if current_comp and len(row_vals) > 0:
            padded = [safe_str(v) for v in row.values]
            new_row = [current_comp] + padded
            processed_rows.append(new_row)

    max_len = max(len(r) for r in processed_rows) if processed_rows else 2
    base_col_names = ['会社名', '車番', '入力番号', '詳細', '備考', '削除対象', '風体', '登録日時', '削除フラグ']
    
    columns = []
    for i in range(max_len):
        if i < len(base_col_names):
            columns.append(base_col_names[i])
        else:
            columns.append(f"extra_{i}")
    
    padded_rows = []
    for r in processed_rows:
        while len(r) < len(columns):
            r.append("")
        padded_rows.append(r)
        
    df = pd.DataFrame(padded_rows, columns=columns[:len(padded_rows[0])] if padded_rows else columns)
    df = df.dropna(subset=['会社名'])
    
    if '削除対象' in df.columns:
        df = df.drop(columns=['削除対象'])
    if '削除フラグ' in df.columns:
        df = df[df['削除フラグ'] != '1']
        
    for c in ['車番', '入力番号', '詳細', '備考', '風体']:
        if c not in df.columns:
            df[c] = ""
            
    for col in df.columns:
        df[col] = df[col].apply(safe_str)
        
    return df

# アプリ全体で共有するデータ
if "app_df" not in st.session_state:
    st.session_state.app_df = load_raw_data()
if "search_box_main" not in st.session_state:
    st.session_state.search_box_main = ""
if "pending_key" not in st.session_state:
    st.session_state.pending_key = None
if "active_card_key" not in st.session_state:
    st.session_state.active_card_key = None
if "scanned_shaban" not in st.session_state:
    st.session_state.scanned_shaban = ""
if "scanned_detail" not in st.session_state:
    st.session_state.scanned_detail = ""
if "action_notice" not in st.session_state:
    st.session_state.action_notice = ""
if "edit_search_keyword" not in st.session_state:
    st.session_state.edit_search_keyword = ""

# テンキー処理
if st.session_state.pending_key is not None:
    if st.session_state.pending_key == "CLEAR":
        st.session_state.search_box_main = ""
    else:
        st.session_state.search_box_main = str(st.session_state.get("search_box_main", "")) + str(st.session_state.pending_key)
    st.session_state.pending_key = None
    st.session_state.active_card_key = None

# データフレームへ会社番号順に新規車両を挿入・ソート
def insert_vehicle_record(company_str, shaban, input_no, detail, remark, fuutai):
    df = st.session_state.app_df.copy()
    
    new_data = {
        '会社名': safe_str(company_str),
        '車番': safe_str(shaban),
        '入力番号': safe_str(input_no),
        '詳細': safe_str(detail),
        '備考': safe_str(remark),
        '風体': safe_str(fuutai),
    }
    for col in df.columns:
        if col not in new_data:
            new_data[col] = ""

    new_row_df = pd.DataFrame([new_data])
    df = pd.concat([df, new_row_df], ignore_index=True)

    def get_comp_code(val):
        m = re.match(r'^(\d{1,2})', safe_str(val))
        return int(m.group(1)) if m else 999

    df['_comp_num'] = df['会社名'].apply(get_comp_code)
    df = df.sort_values(by=['_comp_num', '車番'], ascending=[True, True], kind='stable').drop(columns=['_comp_num']).reset_index(drop=True)
    st.session_state.app_df = df

    try:
        wb = openpyxl.load_workbook(FILE_PATH)
        ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active
        
        comp_found = False
        norm_target = normalize_text(company_str)
        for r in range(1, ws.max_row + 1):
            v = normalize_text(ws.cell(row=r, column=1).value)
            if v == norm_target:
                comp_found = True
                break
        
        if not comp_found:
            r1 = ws.max_row + 1
            ws.cell(row=r1, column=1, value=str(company_str))
            r2 = r1 + 1
            ws.cell(row=r2, column=1, value="")
            ws.cell(row=r2, column=2, value=str(shaban))
            ws.cell(row=r2, column=3, value=str(input_no))
            ws.cell(row=r2, column=4, value=str(detail))