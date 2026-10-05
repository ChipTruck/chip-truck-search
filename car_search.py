import streamlit as st
import pandas as pd

# ページの設定
st.set_page_config(page_title="登録車両検索アプリ", page_icon="🚗", layout="centered")

# メインタイトル
st.title("🚗 登録車両 検索アプリ")
st.write("車両番号や管理番号、会社名などで素早く検索できます。")

# データの読み込み関数
@st.cache_data
def load_data():
    file_path = "新_登録車両資料_連動版.xlsx"
    df = pd.read_excel(file_path, sheet_name="新_登録車両資料")
    return df

try:
    df = load_data()
    
    # 検索ボックス
    search_query = st.text_input("🔍 キーワード検索", "")

    # フィルタリング処理
    if search_query:
        mask = df.astype(str).apply(lambda x: x.str.contains(search_query, case=False, na=False)).any(axis=1)
        filtered_df = df[mask]
    else:
        filtered_df = df

    st.write(f"検索結果: **{len(filtered_df)}** 件")

    # テーブルとして表示
    st.dataframe(filtered_df, use_container_width=True)

    # 詳細確認用のセクション
    st.divider()
    st.subheader("📋 その他のシート情報")
    
    xls = pd.ExcelFile("新_登録車両資料_連動版.xlsx")
    selected_sheet = st.selectbox("確認したいシートを選択", xls.sheet_names)
    
    if selected_sheet:
        sheet_df = pd.read_excel("新_登録車両資料_連動版.xlsx", sheet_name=selected_sheet)
        st.dataframe(sheet_df, use_container_width=True)

except Exception as e:
    st.error(f"データの読み込み中にエラーが発生しました: {e}")