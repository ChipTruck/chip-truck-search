import streamlit as st
import pandas as pd

# ページの設定
st.set_page_config(page_title="登録車両検索アプリ", page_icon="🚗", layout="centered")

# セッション状態の初期化
if "search_query" not in st.session_state:
    st.session_state.search_query = ""

# データの読み込み＆会社名と車両データの整理
@st.cache_data
def load_data():
    file_path = "新_登録車両資料_連動版.xlsx"
    raw_df = pd.read_excel(file_path, sheet_name="新_登録車両資料", header=None)
    raw_df = raw_df.dropna(how="all").astype(str)
    raw_df = raw_df.replace(r'^\s*$', pd.NA, regex=True)

    processed_rows = []
    current_comp = ""
    
    for _, row in raw_df.iterrows():
        row_text = " ".join(row.dropna().astype(str))
        if ":" in row_text or "：" in row_text:
            for val in row.dropna():
                if ":" in str(val) or "：" in str(val):
                    current_comp = str(val).strip()
                    break
            continue 
        
        if current_comp:
            new_row = [current_comp] + list(row.values)
            processed_rows.append(new_row)

    max_len = max(len(r) for r in processed_rows) if processed_rows else 2
    base_col_names = ['会社名', '車番', '入力番号', '詳細', '備考', '削除対象', '風体']
    
    columns = []
    for i in range(max_len):
        if i < len(base_col_names):
            columns.append(base_col_names[i])
        else:
            columns.append(f"extra_{i}")
    
    padded_rows = [r + [pd.NA] * (max_len - len(r)) for r in processed_rows]
    df_display = pd.DataFrame(padded_rows, columns=columns)
    df_display = df_display.dropna(subset=['会社名'])
    
    if '削除対象' in df_display.columns:
        df_display = df_display.drop(columns=['削除対象'])
    
    return df_display

try:
    df_base = load_data()

    # ── ヘルパー機能：タイトル ＆ ホームに戻るボタン ──
    col_title, col_home = st.columns([4, 1])
    with col_title:
        st.title("🚗 登録車両 検索アプリ")
    with col_home:
        st.write("") 
        if st.button("🏠 ホーム", use_container_width=True):
            st.session_state.search_query = ""
            st.rerun()

    st.write("車番（ナンバープレートの数字など）で素早く検索できます。")

    # 通常のテキスト入力欄（キーボード入力用）
    def update_from_input():
        st.session_state.search_query = st.session_state.temp_input

    search_query = st.text_input(
        "🔍 車番を入力（例: 1, 8, 14 など）", 
        value=st.session_state.search_query,
        key="temp_input",
        on_change=update_from_input
    )

    # ── 画面上のテンキーボタンエリア ──
    st.write("🔢 **テンキー入力ボタン**")
    
    r1_c1, r1_c2, r1_c3 = st.columns(3)
    if r1_c1.button("1", use_container_width=True): st.session_state.search_query += "1"; st.rerun()
    if r1_c2.button("2", use_container_width=True): st.session_state.search_query += "2"; st.rerun()
    if r1_c3.button("3", use_container_width=True): st.session_state.search_query += "3"; st.rerun()

    r2_c1, r2_c2, r2_c3 = st.columns(3)
    if r2_c1.button("4", use_container_width=True): st.session_state.search_query += "4"; st.rerun()
    if r2_c2.button("5", use_container_width=True): st.session_state.search_query += "5"; st.rerun()
    if r2_c3.button("6", use_container_width=True): st.session_state.search_query += "6"; st.rerun()

    r3_c1, r3_c2, r3_c3 = st.columns(3)
    if r3_c1.button("7", use_container_width=True): st.session_state.search_query += "7"; st.rerun()
    if r3_c2.button("8", use_container_width=True): st.session_state.search_query += "8"; st.rerun()
    if r3_c3.button("9", use_container_width=True): st.session_state.search_query += "9"; st.rerun()

    r4_c1, r4_c2, r4_c3 = st.columns(3)
    if r4_c1.button("0", use_container_width=True): st.session_state.search_query += "0"; st.rerun()
    if r4_c2.button("⌫ 1文字消す", use_container_width=True): 
        st.session_state.search_query = st.session_state.search_query[:-1]
        st.rerun()
    if r4_c3.button("クリア", use_container_width=True): 
        st.session_state.search_query = ""
        st.rerun()

    st.divider()

    # フィルタリング処理（【車番】の列だけで完全ピンポイント検索）
    filtered_df = df_base.copy()
    is_searched = False

    if st.session_state.search_query:
        is_searched = True
        if '車番' in filtered_df.columns:
            mask = filtered_df['車番'].str.contains(st.session_state.search_query, case=False, na=False)
            filtered_df = filtered_df[mask]
        else:
            mask = filtered_df.apply(lambda x: x.str.contains(st.session_state.search_query, case=False, na=False)).any(axis=1)
            filtered_df = filtered_df[mask]

    # 検索されたときだけ、数字の小さい順に正確に並び替える
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

    # ── 表示用：同じ会社名が連続する場合は最初だけ表示して、下は空白にする ──
    display_df = filtered_df.copy()
    display_df['会社名'] = display_df['会社名'].mask(display_df['会社名'] == display_df['会社名'].shift(), '')

    if is_searched:
        st.write(f"検索結果: **{len(filtered_df)}** 行（数字の小さい順に表示中）")
    else:
        st.write(f"全車両データ一覧: **{len(filtered_df)}** 行")

    # 表として表示
    st.dataframe(display_df, width="stretch", hide_index=True)

    # 詳細確認用のセクション（その他のシート）
    st.divider()
    st.subheader("📋 その他のシート情報")
    
    xls = pd.ExcelFile("新_登録車両資料_連動版.xlsx")
    selected_sheet = st.selectbox("確認したいシートを選択", xls.sheet_names)
    
    if selected_sheet:
        sheet_df = pd.read_excel("新_登録車両資料_連動版.xlsx", sheet_name=selected_sheet, header=None)
        sheet_df = sheet_df.dropna(how="all").astype(str)
        sheet_df = sheet_df.replace(r'^\s*$', pd.NA, regex=True)
        st.dataframe(sheet_df, width="stretch", hide_index=True)

except Exception as e:
    st.error(f"データの読み込み中にエラーが発生しました: {e}")