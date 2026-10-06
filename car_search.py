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
            ws.cell(row=r2, column=5, value=str(remark))
            ws.cell(row=r2, column=7, value=str(fuutai))
        else:
            r = ws.max_row + 1
            ws.cell(row=r, column=1, value="")
            ws.cell(row=r, column=2, value=str(shaban))
            ws.cell(row=r, column=3, value=str(input_no))
            ws.cell(row=r, column=4, value=str(detail))
            ws.cell(row=r, column=5, value=str(remark))
            ws.cell(row=r, column=7, value=str(fuutai))
            
        wb.save(FILE_PATH)
    except Exception:
        pass

# 車両の完全削除
def delete_vehicle_record(company_str, old_shaban, old_input_no):
    df = st.session_state.app_df.copy()
    norm_c = normalize_text(company_str)
    norm_s = normalize_text(old_shaban)
    norm_inp = normalize_text(old_input_no)

    drop_indices = []
    for idx, r in df.iterrows():
        r_c = normalize_text(r.get('会社名', ''))
        r_s = normalize_text(r.get('車番', ''))
        r_inp = normalize_text(r.get('入力番号', ''))
        if (norm_inp and r_inp == norm_inp) or (r_c == norm_c and (r_s == norm_s or (not norm_s and not r_s))):
            drop_indices.append(idx)

    if drop_indices:
        df = df.drop(index=drop_indices).reset_index(drop=True)
        st.session_state.app_df = df

    try:
        wb = openpyxl.load_workbook(FILE_PATH)
        ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active
        
        current_comp = ""
        rows_to_delete = []
        for r in range(1, ws.max_row + 1):
            c1 = normalize_text(ws.cell(row=r, column=1).value)
            if ":" in c1:
                current_comp = c1
            c2 = normalize_text(ws.cell(row=r, column=2).value)
            c3 = normalize_text(ws.cell(row=r, column=3).value)
            if (norm_inp and c3 == norm_inp) or (current_comp == norm_c and (c2 == norm_s or (not norm_s and not c2))):
                rows_to_delete.append(r)
        
        for r in reversed(rows_to_delete):
            ws.delete_rows(r)
            
        wb.save(FILE_PATH)
    except Exception:
        pass

# 車両情報の更新
def update_vehicle_record(old_company, old_shaban, old_input_no, new_comp, new_shaban, new_input_no, new_detail, new_remark, new_fuutai):
    df = st.session_state.app_df.copy()
    norm_c = normalize_text(old_company)
    norm_s = normalize_text(old_shaban)
    norm_inp = normalize_text(old_input_no)

    for idx, r in df.iterrows():
        r_c = normalize_text(r.get('会社名', ''))
        r_s = normalize_text(r.get('車番', ''))
        r_inp = normalize_text(r.get('入力番号', ''))
        
        if (norm_inp and r_inp == norm_inp) or (r_c == norm_c and (r_s == norm_s or (not norm_s and not r_s))):
            df.at[idx, '会社名'] = safe_str(new_comp)
            df.at[idx, '車番'] = safe_str(new_shaban)
            df.at[idx, '入力番号'] = safe_str(new_input_no)
            df.at[idx, '詳細'] = safe_str(new_detail)
            df.at[idx, '備考'] = safe_str(new_remark)
            df.at[idx, '風体'] = safe_str(new_fuutai)
            break
            
    st.session_state.app_df = df

    try:
        wb = openpyxl.load_workbook(FILE_PATH)
        ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active
        
        current_comp = ""
        for r in range(1, ws.max_row + 1):
            c1 = normalize_text(ws.cell(row=r, column=1).value)
            if ":" in c1:
                current_comp = c1
            c2 = normalize_text(ws.cell(row=r, column=2).value)
            c3 = normalize_text(ws.cell(row=r, column=3).value)

            if (norm_inp and c3 == norm_inp) or (current_comp == norm_c and (c2 == norm_s or (not norm_s and not c2))):
                ws.cell(row=r, column=2, value=str(new_shaban))
                ws.cell(row=r, column=3, value=str(new_input_no))
                ws.cell(row=r, column=4, value=str(new_detail))
                ws.cell(row=r, column=5, value=str(new_remark))
                ws.cell(row=r, column=7, value=str(new_fuutai))
                break
        wb.save(FILE_PATH)
    except Exception:
        pass

# Upload.csv 生成関数
def generate_upload_csv(df_source):
    rows = []
    for _, r in df_source.iterrows():
        v_id = safe_str(r.get('入力番号', ''))
        fuutai_raw = safe_str(r.get('風体', ''))
        
        if v_id:
            m_fuutai = re.findall(r'\d+', fuutai_raw)
            fuutai_val = m_fuutai[0] if m_fuutai else "0"
            rows.append({
                'Vehicle ID': v_id,
                'Max weight': 0,
                'Weight': fuutai_val
            })
            
    csv_df = pd.DataFrame(rows, columns=['Vehicle ID', 'Max weight', 'Weight'])
    return csv_df.to_csv(index=False, encoding='utf-8-sig')

df_base = st.session_state.app_df

# ── ヘッダー：タイトル ＆ ホーム ＆ Upload.csvダウンロードボタン ──
col_title, col_home, col_dl = st.columns([2.0, 0.9, 1.1])
with col_title:
    st.title("🚗 車両管理＆検索")
with col_home:
    st.write("")
    if st.button("🏠 ホーム", use_container_width=True):
        st.session_state.search_box_main = ""
        st.session_state.active_card_key = None
        st.session_state.scanned_shaban = ""
        st.session_state.scanned_detail = ""
        st.session_state.action_notice = ""
        st.rerun()
with col_dl:
    st.write("")
    csv_data = generate_upload_csv(df_base)
    st.download_button(
        label="📥 Upload.csv",
        data=csv_data,
        file_name="Upload.csv",
        mime="text/csv",
        use_container_width=True
    )

tab1, tab2, tab3 = st.tabs(["🔍 検索", "➕ 新規登録", "✏️ 編集・削除"])

# ── 【タブ1】 検索・閲覧 ──
with tab1:
    st.write("車番の数字で素早く検索できます。")

    st.text_input(
        "🔍 車番を入力（例: 1, 8, 14 など）", 
        key="search_box_main"
    )

    def on_num_click(val):