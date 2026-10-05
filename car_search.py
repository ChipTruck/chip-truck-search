import streamlit as st
import pandas as pd

# ページの設定
st.set_page_config(page_title="登録車両検索アプリ", page_icon="🚗", layout="centered")

# データの読み込み＆会社名だけを正確に引き継ぐ関数
@st.cache_data
def load_data():
    file_path = "新_登録車両資料_連動版.xlsx"
    df = pd.read_excel(file_path, sheet_name="新_登録車両資料", header=None)
    
    # 完全な空行を削除
    df = df.dropna(how="all")
    
    # 文字列に変換して空白を整理
    df = df.astype(str)
    df = df.replace(r'^\s*$', pd.NA, regex=True)
    
    # 会社名が入っている列（通常は0列目または1列目）を特定して、会社名だけを下に引き継ぐ
    # 車両番号の列には絶対に数字が流れ込まないように独立させる
    company_col_idx = 0
    for idx, col in enumerate(df.columns):
        # 「:」や「：」が含まれているセルが多い列を会社名列とみなす
        if df[col].astype(str).str.contains('[:：]', na=False).sum() > 0:
            company_col_idx = col
            break

    # 会社名専用の新しい列を作る
    df['company_name'] = df[company_col_idx]
    df['company_name'] = df['company_name'].ffill() # 会社名だけを下に流す
    
    return df

try:
    df = load_data()

    # 会社名列を一番左に配置して、元の会社名があった列は削除する
    cols = ['company_name'] + [col for col in df.columns if col not in ['company_name', 0]]
    df_display = df[cols]

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
        if len(filtered_df.columns) > 2:
            # 会社名列を除外した車両データ側の列だけで検索
            target_cols = filtered_df.columns[1:]
            mask = filtered_df[target_cols].apply(lambda x: x.str.contains(search_query, case=False, na=False)).any(axis=1)
            filtered_df = filtered_df[mask]
        else:
            mask = filtered_df.apply(lambda x: x.str.contains(search_query, case=False, na=False)).any(axis=1)
            filtered_df = filtered_df[mask]

    # 検索されたときだけ、数字の小さい順に正確に並び替える
    if is_searched:
        try:
            if len(filtered_df.columns) > 1:
                sort_col = filtered_df.columns[1] # 車両番号の列
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