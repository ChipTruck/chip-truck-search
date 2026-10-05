import streamlit as st
import pandas as pd

# ページの設定
st.set_page_config(page_title="登録車両検索アプリ", page_icon="🚗", layout="centered")

# データの読み込み＆会社名と車両データを完全に分離・整理する関数
@st.cache_data
def load_data():
    file_path = "新_登録車両資料_連動版.xlsx"
    df = pd.read_excel(file_path, sheet_name="新_登録車両資料", header=None)
    
    # 完全な空行を削除
    df = df.dropna(how="all")
    
    # 文字列に変換して空白を整理
    df = df.astype(str)
    df = df.replace(r'^\s*$', pd.NA, regex=True)
    
    # 1. 会社名（「:」や「：」を含む行、または特定のパターン）を判定して新しい「company」列を作る
    company_list = []
    current_company = "未分類"
    
    for idx, row in df.iterrows():
        # 行の中に「:」や「：」が含まれているセルがあるかチェック
        row_str = " ".join(row.dropna().astype(str))
        if ":" in row_str or "：" in row_str:
            # 会社名らしい文字列を取得
            for val in row.dropna():
                val_s = str(val).strip()
                if ":" in val_s or "：" in val_s:
                    current_company = val_s
                    break
        company_list.append(current_company)
        
    df.insert(0, 'company_name', company_list)
    
    # 2. 会社名行自体のデータ（車両が入っていない行）は、表の見た目をスッキリさせるために除外、
    #    または車両データ側に正しく配置する
    # 会社名文字列そのものが含まれている行（車両データがない行）を特定して削除
    is_header_row = df.apply(lambda row: row.astype(str).str.contains('[:：]').any(), axis=1)
    # ただし車両データにもコロンが含まれる可能性を考慮し、主要なセルが空っぽの行を会社名行とみなす
    df_cleaned = df[~is_header_row].copy()
    
    # もし会社名だけの行も残したい場合はそのままですが、今回は「会社名ごとにグループ化して右側にデータを出す」ため、
    # 会社名行から車両データを分離します
    
    return df

try:
    # 実際にすっきりとデータを再構築するロジック
    file_path = "新_登録車両資料_連動版.xlsx"
    raw_df = pd.read_excel(file_path, sheet_name="新_登録車両資料", header=None)
    raw_df = raw_df.dropna(how="all").astype(str)
    raw_df = raw_df.replace(r'^\s*$', pd.NA, regex=True)

    # 会社名を下に引き継ぎつつ、会社名行の不要なズレを解消する処理
    processed_rows = []
    current_comp = ""
    
    for _, row in raw_df.iterrows():
        row_text = " ".join(row.dropna().astype(str))
        # 会社名の行か判定
        if ":" in row_text or "：" in row_text:
            for val in row.dropna():
                if ":" in str(val) or "：" in str(val):
                    current_comp = str(val).strip()
                    break
            continue # 会社名行そのものはデータ行とは別にするためスキップ
        
        # 車両データの行の場合
        if current_comp:
            new_row = [current_comp] + list(row.values)
            processed_rows.append(new_row)

    # 新しいきれいなデータフレームを作成
    max_len = max(len(r) for r in processed_rows) if processed_rows else 2
    columns = ['会社名'] + [f"col_{i}" for i in range(max_len - 1)]
    
    # 行の長さを合わせる
    padded_rows = [r + [pd.NA] * (max_len - len(r)) for r in processed_rows]
    df_display = pd.DataFrame(padded_rows, columns=columns)
    df_display = df_display.dropna(subset=['会社名'])

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
    filtered_df = df_display.copy()
    is_searched = False

    if search_query:
        is_searched = True
        if len(filtered_df.columns) > 1:
            # 会社名列（0番目）を除外した右側の車両データ列だけで検索
            target_cols = filtered_df.columns[1:]
            mask = filtered_df[target_cols].apply(lambda x: x.str.contains(search_query, case=False, na=False)).any(axis=1)
            filtered_df = filtered_df[mask]
        else:
            mask = filtered_df.apply(lambda x: x.str.contains(search_query, case=False, na=False)).any(axis=1)
            filtered_df = filtered_df[mask]

    # 検索されたときだけ、数字の小さい順に正確に並び替える
    if is_searched:
        try:
            if len(filtered_df.columns) > 2:
                sort_col = filtered_df.columns[1] # 車両の数字列
                filtered_df = filtered_df.copy()
                filtered_df['_sort_num'] = pd.to_numeric(filtered_df[sort_col].str.extract(r'(\d+)', expand=False), errors='coerce')
                filtered_df = filtered_df.sort_values(by='_sort_num', ascending=True, na_position='last')
                filtered_df = filtered_df.drop(columns=['_sort_num'])
        except Exception:
            pass

    if is_searched:
        st.write(f"検索結果: **{len(filtered_df)}** 行（数字の小さい順に表示中）")
    else:
        st.write(f"全車両データ一覧: **{len(filtered_df)}** 行")

    # 表として表示
    st.dataframe(filtered_df, width="stretch", hide_index=True)

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