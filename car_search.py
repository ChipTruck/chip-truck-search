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

# ── 見た目の調整＆鎖マーク完全消去CSS ──
st.markdown("""
<style>
.app-main-title {
    font-size: 1.55rem !important;
    font-weight: bold !important;
    color: #1e88e5;
    margin-top: -10px;
    margin-bottom: 8px;
    white-space: nowrap !important;
}

[data-testid="stDataFrame"] {
    width: 100% !important;
}

a.anchor-link, [data-testid="stHeaderActionElements"] {
    display: none !important;
}

/* 検索ボックスの文字サイズと高さ */
div[data-testid="stTextInput"] input {
    font-size: 1.4rem !important;
    height: 50px !important;
    font-weight: bold !important;
    border: 2px solid #1e88e5 !important;
    border-radius: 8px !important;
    padding: 0 12px !important;
}

div[data-testid="stForm"] {
    border: none !important;
    padding: 0 !important;
}

div[data-testid="stHorizontalBlock"] {
    display: flex !important;
    flex-direction: row !important;
    flex-wrap: nowrap !important;
    align-items: center !important;
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

def safe_str(val):
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    if s.lower() in ["none", "nan", "<na>"]:
        return ""
    return s

def normalize_text(text):
    s = safe_str(text)
    if not s:
        return ""
    t = s.replace("：", ":").replace(" ", "").replace(" ", "")
    return t.strip()

# ── 常にPCと共通のExcel実ファイルからデータを読み込む ──
@st.cache_data(ttl=5)
def load_shared_excel_data():
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
        
    def get_comp_code(val):
        m = re.match(r'^(\d{1,2})', safe_str(val))
        return int(m.group(1)) if m else 999

    df['_comp_num'] = df['会社名'].apply(get_comp_code)
    df = df.sort_values(by=['_comp_num', '車番'], ascending=[True, True], kind='stable').drop(columns=['_comp_num']).reset_index(drop=True)
    return df

# セッション状態の管理
if "search_query" not in st.session_state:
    st.session_state.search_query = ""
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
if "form_reset_counter" not in st.session_state:
    st.session_state.form_reset_counter = 0

def reset_to_home():
    st.session_state.search_query = ""
    st.session_state.active_card_key = None
    st.session_state.action_notice = ""
    st.session_state.form_reset_counter += 1

def insert_vehicle_record(company_str, shaban, input_no, detail, remark, fuutai):
    wb = openpyxl.load_workbook(FILE_PATH)
    ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active
    
    comp_row_idx = None
    target_comp_code = -1
    m_target = re.match(r'^(\d{1,2})[:：]', company_str.strip())
    if m_target:
        target_comp_code = int(m_target.group(1))

    insert_before_row = None
    for r in range(1, ws.max_row + 1):
        v = normalize_text(ws.cell(row=r, column=1).value)
        if not v:
            continue
        if normalize_text(company_str) == v:
            comp_row_idx = r
            break
        m = re.match(r'^(\d{1,2})[:：]', v)
        if m and target_comp_code >= 0:
            c_code = int(m.group(1))
            if c_code > target_comp_code and insert_before_row is None:
                insert_before_row = r

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if comp_row_idx is not None:
        insert_target = comp_row_idx + 1
        while insert_target <= ws.max_row:
            cell_v = str(ws.cell(row=insert_target, column=1).value or "").strip()
            if (":" in cell_v or "：" in cell_v) and normalize_text(cell_v) != normalize_text(company_str):
                break
            insert_target += 1
        
        ws.insert_rows(insert_target)
        ws.cell(row=insert_target, column=1, value="")
        ws.cell(row=insert_target, column=2, value=str(shaban))
        ws.cell(row=insert_target, column=3, value=str(input_no))
        ws.cell(row=insert_target, column=4, value=str(detail))
        ws.cell(row=insert_target, column=5, value=str(remark))
        ws.cell(row=insert_target, column=6, value="")
        ws.cell(row=insert_target, column=7, value=str(fuutai))
        ws.cell(row=insert_target, column=8, value=now_str)
    else:
        ins_r = insert_before_row if insert_before_row is not None else (ws.max_row + 1)
        ws.insert_rows(ins_r, amount=2)
        ws.cell(row=ins_r, column=1, value=str(company_str))
        ws.cell(row=ins_r + 1, column=1, value="")
        ws.cell(row=ins_r + 1, column=2, value=str(shaban))
        ws.cell(row=ins_r + 1, column=3, value=str(input_no))
        ws.cell(row=ins_r + 1, column=4, value=str(detail))
        ws.cell(row=ins_r + 1, column=5, value=str(remark))
        ws.cell(row=ins_r + 1, column=6, value="")
        ws.cell(row=ins_r + 1, column=7, value=str(fuutai))
        ws.cell(row=ins_r + 1, column=8, value=now_str)
        
    wb.save(FILE_PATH)
    st.cache_data.clear()

def delete_vehicle_record(company_str, old_shaban, old_input_no):
    wb = openpyxl.load_workbook(FILE_PATH)
    ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active
    norm_c = normalize_text(company_str)
    norm_s = normalize_text(old_shaban)
    norm_inp = normalize_text(old_input_no)

    current_comp = ""
    rows_to_delete = []
    for r in range(1, ws.max_row + 1):
        c1 = normalize_text(ws.cell(row=r, column=1).value)
        if ":" in c1:
            current_comp = c1
        c2 = normalize_text(ws.cell(row=r, column=2).value)
        c3 = normalize_text(ws.cell(row=r, column=3).