import streamlit as st
import pandas as pd

# ページの設定
st.set_page_config(page_title="登録車両検索アプリ", page_icon="🚗", layout="centered")

# メインタイトル
st.title("🚗 登録車両 検索アプリ")
st.write("車両番号や管理番号、会社名などで素早く検索できます。")

# データの読み込み＆強力なデータクリーニング関数
@st.cache_data
def load_data():
    file_path = "新_登録車両資料_連動版.xlsx"
    df = pd.read_excel(file_path, sheet_name="新_登録車両資料")
    
    # 1. すべてのセルが完全に空の行を削除
    df = df.dropna(how="all")
    
    # 2. 一度文字列に変換してから、スペースやタブだけの空行をNaNに置換
    df = df.astype(str)
    df = df.replace(r'^\s*$', pd.NA, regex=True)
    
    # 3. 会社名や項目が入っている列（2列目）を基準に、空っぽやNoneの行を徹底的に除外
    if len(df.columns) > 1:
        target_col = df.columns[1]
        df = df.dropna(subset=[target_col])
        df = df[~df[target_col].isin(["None", "nan", "nat", "None.0", "nan.0", "<NA>"])]
        
    return df

try:
    df = load_data()
    target_col = df.columns[1] if len(df.columns) > 1 else df.columns[0]
    
    # ── 会社名ごとの絞り込みセレクトボックス（画面上部に配置） ──
    st.subheader("🏢 会社名で絞り込み")
    categories = ["すべて表示"] + list(df[target_col].unique())
    # ユーザーがパッと選んで一覧表示できるようにする
    selected_category = st.selectbox("表示したい会社名を選択してください", categories)

    # 検索ボックス
    search_query = st.text_input("🔍 キーワード検索（車両番号など）", "")

    # フィルタリング処理
    filtered_df = df
    if selected_category != "すべて表示":
        filtered_df = filtered_df[filtered_df[target_col] == selected_category]

    if search_query:
        mask = filtered_df.apply(lambda x: x.str.contains(search_query, case=False, na=False)).any(axis=1)
        filtered_df = filtered_df[mask]

    st.write(f"検索結果: **{len(filtered_df)}** 件")

    # テーブルとして表示
    st.dataframe(filtered_df, width="stretch")

    # 詳細確認用のセクション（その他のシート）
    st.divider()
    st.subheader("📋 その他のシート情報")
    
    xls = pd.ExcelFile("新_登録車両資料_連動版.xlsx")
    selected_sheet = st.selectbox("確認したいシートを選択", xls.sheet_names)
    
    if selected_sheet:
        sheet_df = pd.read_excel("新_登録車両資料_連動版.xlsx", sheet_name=selected_sheet)
        sheet_df = sheet_df.dropna(how="all")
        sheet_df = sheet_df.astype(str)
        sheet_df = sheet_df.replace(r'^\s*$', pd.NA, regex=True)
        if len(sheet_df.columns) > 1:
            s_col = sheet_df.columns[1]
            sheet_df = sheet_df.dropna(subset=[s_col])
            sheet_df = sheet_df[~sheet_df[s_col].isin(["None", "nan", "nat", "None.0", "nan.0", "<NA>"])]
        st.dataframe(sheet_df, width="stretch")

except Exception as e:
    st.error(f"データの読み込み中にエラーが発生しました: {e}")