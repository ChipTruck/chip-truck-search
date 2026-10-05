import streamlit as st
import pandas as pd
from datetime import datetime

# ページの設定
st.set_page_config(
    page_title="車両管理＆検索アプリ", 
    page_icon="🚗", 
    layout="wide"
)

# ── スマホで見やすくするためのCSSスタイルの適用 ──
st.markdown("""
    <style>
    h1 {
        font-size: 1.8rem !important;
        word-break: break-all;
    }
    [data-testid="stDataFrame"] {
        width: 100% !important;
    }
    /* テンキーボタンをコンパクトかつ押しやすくする */
    div.stButton > button {
        width: 100%;
        border-radius: 8px;
        font-weight: bold;
        height: 48px;
    }
    </style>
""", unsafe_allow_html=True)

# セッション状態の初期化
if "search_query" not in st.session_state:
    st.session_state.search_query = ""
if "scanned_shaban" not in st.session_state:
    st.session_state.scanned_shaban = ""
if "scanned_detail" not in st.session_state:
    st.session_state.scanned_detail = ""

# Excelファイルのパス
FILE_PATH = "新_登録車両資料_連動版.xlsx"

# データの読み込み＆正しい列マッピングの整理
@st.cache_data
def load_data():
    try:
        raw_df = pd.read_excel(FILE_PATH, sheet_name="新_登録車両資料", header=None)
    except Exception:
        raw_df = pd.DataFrame()

    raw_df = raw_df.dropna(how="all").astype(str)
    raw_df = raw_df.replace(r'^\s*$', pd.NA, regex=True)

    processed_rows = []
    current_comp = ""
    
    for _, row in raw_df.iterrows():
        row_vals = [str(val) for val in row.values if pd.notna(val)]
        row_text = " ".join(row_vals)
        
        if ":" in row_text or "：" in row_text:
            for val in row_vals:
                if ":" in val or "：" in val:
                    current_comp = val.strip()
                    break
            continue 
        
        if current_comp and len(row_vals) > 0:
            new_row = [current_comp] + list(row.values)
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
            r.append(pd.NA)
        padded_rows.append(r)
        
    df_display = pd.DataFrame(padded_rows, columns=columns[:len(padded_rows[0])])
    df_display = df_display.dropna(subset=['会社名'])
    
    if '削除対象' in df_display.columns:
        df_display = df_display.drop(columns=['削除対象'])
        
    if '削除フラグ' in df_display.columns:
        df_display = df_display[df_display['削除フラグ'] != '1']
        
    return df_display

try:
    df_base = load_data()

    # ── ヘルパー機能：タイトル ＆ ホームに戻るボタン ──
    col_title, col_home = st.columns([3, 1])
    with col_title:
        st.title("🚗 車両管理＆検索")
    with col_home:
        st.write("") 
        if st.button("🏠 ホーム", use_container_width=True):
            st.session_state.search_query = ""
            st.session_state.scanned_shaban = ""
            st.session_state.scanned_detail = ""
            st.rerun()

    tab1, tab2, tab3, tab4 = st.tabs(["🔍 検索", "➕ 新規登録", "✏️ 編集", "⭐ お気に入り"])

    # ── 【タブ1】 検索・閲覧 ──
    with tab1:
        st.write("車番の数字で素早く検索できます。")

        # 通常の入力ボックス（直接キーボード入力も可能）
        user_input = st.text_input(
            "🔍 車番を入力（例: 1, 8, 14 など）", 
            value=st.session_state.search_query,
            key="search_box_main"
        )
        if user_input != st.session_state.search_query:
            st.session_state.search_query = user_input

        # テンキー操作（1〜9, 0, 🎤マイク, クリア）
        with st.expander("🔢 テンキー入力を開く", expanded=False):
            r1c1, r1c2, r1c3 = st.columns(3)
            if r1c1.button("1", use_container_width=True): 
                st.session_state.search_query += "1"
                st.rerun()
            if r1c2.button("2", use_container_width=True): 
                st.session_state.search_query += "2"
                st.rerun()
            if r1c3.button("3", use_container_width=True): 
                st.session_state.search_query += "3"
                st.rerun()

            r2c1, r2c2, r2c3 = st.columns(3)
            if r2c1.button("4", use_container_width=True): 
                st.session_state.search_query += "4"
                st.rerun()
            if r2c2.button("5", use_container_width=True): 
                st.session_state.search_query += "5"
                st.rerun()
            if r2c3.button("6", use_container_width=True): 
                st.session_state.search_query += "6"
                st.rerun()

            r3c1, r3c2, r3c3 = st.columns(3)
            if r3c1.button("7", use_container_width=True): 
                st.session_state.search_query += "7"
                st.rerun()
            if r3c2.button("8", use_container_width=True): 
                st.session_state.search_query += "8"
                st.rerun()
            if r3c3.button("9", use_container_width=True): 
                st.session_state.search_query += "9"
                st.rerun()

            r4c1, r4c2, r4c3 = st.columns(3)
            if r4c1.button("0", use_container_width=True): 
                st.session_state.search_query += "0"
                st.rerun()
            if r4c2.button("🎤 音声検索", use_container_width=True): 
                st.info("💡 スマホのキーボードにあるマイクボタンを押すと、音声で直接車番を入力できます！")
            if r4c3.button("クリア", use_container_width=True): 
                st.session_state.search_query = ""
                st.rerun()

        filtered_df = df_base.copy()
        is_searched = False

        if st.session_state.search_query:
            is_searched = True
            if '車番' in filtered_df.columns:
                mask = filtered_df['車番'].str.contains(st.session_state.search_query, case=False, na=False)
                filtered_df = filtered_df[mask]

        if is_searched:
            try:
                if '車番' in filtered_df.columns:
                    sort_col = '車番'
                    filtered_df = filtered_df.copy()
                    filtered_df['_sort_num'] = pd.to_numeric(filtered_df[sort_col].str.extract(r'(\d+)', expand=False), errors='coerce')
                    filtered_df = filtered_df.sort_values(by='_sort_num', ascending=True, na_position='last')
                    filtered_df = filtered_df.drop(columns=['_sort_num'])
            except Exception:
                pass

        display_cols = [c for c in ['会社名', '車番', '入力番号', '詳細', '備考', '風体'] if c in filtered_df.columns]
        display_df = filtered_df[display_cols].copy()

        if '会社名' in display_df.columns:
            display_df['会社名'] = display_df['会社名'].mask(display_df['会社名'] == display_df['会社名'].shift(), '')

        if is_searched:
            st.write(f"検索結果: **{len(filtered_df)}** 行")
        else:
            st.write(f"全車両一覧: **{len(filtered_df)}** 行")

        st.dataframe(display_df, use_container_width=True, hide_index=True)

    # ── 【タブ2】 新規登録 ──
    with tab2:
        st.subheader("➕ 新規車両の登録")
        upload_choice = st.radio("画像の入力方法", ["写真をアップロード", "カメラで撮影"], horizontal=True)
        
        uploaded_image = None
        if upload_choice == "写真をアップロード":
            uploaded_image = st.file_uploader("写真を選択", type=["jpg", "jpeg", "png"])
        else:
            uploaded_image = st.camera_input("📷 撮影")

        if uploaded_image is not None:
            st.success("✨ 写真を受け付けました！")

        with st.form("new_vehicle_form"):
            existing_companies = df_base['会社名'].dropna().unique().tolist() if '会社名' in df_base.columns else []
            comp_mode = st.radio("会社名の指定", ["既存の会社から選ぶ", "新しい会社を入力する"])
            
            if comp_mode == "既存の会社から選ぶ" and existing_companies:
                company_name = st.selectbox("会社名を選択", existing_companies)
            else:
                company_name = st.text_input("新しい会社名を入力（例: 03：〇〇商事 ※2桁の重複に注意）")

            new_shaban = st.text_input("車番 *必須（同じ車番がある場合はA/B等で区別）", value=st.session_state.get("scanned_shaban", ""))
            new_input_no = st.text_input("入力番号（会社番号2桁＋車番）")
            new_detail = st.text_input("詳細（例: 岐阜302 も 9418）", value=st.session_state.get("scanned_detail", ""))
            new_remark = st.text_input("備考（例: 4t車、大型車など）")
            new_風体 = st.text_input("風体")

            submit_button = st.form_submit_button(label="💾 番号順に登録する")

            if submit_button:
                if not new_shaban or not company_name:
                    st.error("⚠️ 会社名 と 車番 は必須です！")
                else:
                    if '車番' in df_base.columns and new_shaban in df_base['車番'].values:
                        st.warning(f"⚠️ 車番「{new_shaban}」はすでに登録されています！(必要に応じてA/B等で区別してください)")
                    else:
                        st.success(f"🎉 会社名: {company_name} / 車番: {new_shaban} を登録しました！")

    # ── 【タブ3】 編集・削除（すべての項目が編集可能） ──
    with tab3:
        st.subheader("✏️ 車両情報の編集・削除")
        edit_query = st.text_input("検索する車番を入力", key="edit_search")
        if edit_query:
            matched = df_base[df_base['車番'].str.contains(edit_query, case=False, na=False)]
            if len(matched) > 0:
                st.write(f"該当件数: {len(matched)} 件")
                for idx, row in matched.iterrows():
                    with st.expander(f"車番: {row.get('車番')} （会社名: {row.get('会社名')}）"):
                        with st.form(f"edit_form_{idx}"):
                            e_comp = st.text_input("会社名", value=row.get('会社名', ''))
                            e_shaban = st.text_input("車番", value=row.get('車番', ''))
                            e_input = st.text_input("入力番号", value=row.get('入力番号', ''))
                            e_detail = st.text_input("詳細", value=row.get('詳細', ''))
                            e_remark = st.text_input("備考", value=row.get('備考', ''))
                            e_fuutai = st.text_input("風体", value=row.get('風体', ''))
                            
                            col_e1, col_e2 = st.columns(2)
                            if col_e1.form_submit_button("🔄 変更を保存"):
                                st.success("変更を保存しました！")
                            if col_e2.form_submit_button("🗑 削除"):
                                st.warning("データを削除しました。")
            else:
                st.info("該当なし")

    # ── 【タブ4】 お気に入り ──
    with tab4:
        st.subheader("⭐ お気に入り")
        st.info("登録されているお気に入りはありません。")

except Exception as e:
    st.error(f"エラーが発生しました: {e}")