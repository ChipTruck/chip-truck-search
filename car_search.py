import streamlit as st
import pandas as pd

# ページの設定
st.set_page_config(page_title="登録車両検索アプリ", page_icon="🚗", layout="centered")

# データの読み込み＆列名の整理
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
    
    # 列名の割り当て（指定された通りに設定）
    base_col_names = ['会社名', '車番', '入力番号', '詳細', '備考', '削除対象', '風体']
    
    # 実際のデータの長さに合わせて列名リストを調整
    columns = []
    for i in range(max_len):
        if i < len(base_col_names):
            columns.append(base_col_names[i])
        else:
            columns.append(f"extra_{i}")
    
    padded_rows = [r + [pd.NA] * (max_len - len(r)) for r in processed_rows]
    df_display = pd.DataFrame(padded_rows, columns=columns)
    df_display = df_display.dropna(subset=['会社名'])
    
    # col_4（削除対象）が存在する場合は削除する
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
            st.rerun()

    st.write("車両番号（ナンバープレートの数字）で素早く検索できます。")

    # 検索ボックス
    search_query = st.text_input("🔍 ナンバープレートの数字を入力（例: 1, 1351 など）", "")

    # フィルタリング処理
    filtered_df = df_base.copy()
    is_searched = False

    if search_query:
        is_searched = True
        if len(filtered_df.columns) > 1:
            # 会社名列を除外したデータ側の列だけで検索
            target_cols = filtered_df.columns[1:]
            mask = filtered_df[target_cols].apply(lambda x: x.str.contains(search_query, case=False, na=False)).any(axis=1)
            filtered_df = filtered_df[mask]
        else:
            mask = filtered_df.apply(lambda x: x.str.contains(search_query, case=False, na=False)).any(axis=1)
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