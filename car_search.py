import streamlit as st
import pandas as pd

# ページの設定
st.set_page_config(page_title="登録車両検索アプリ", page_icon="🚗", layout="centered")

# メインタイトル
st.title("🚗 登録車両 検索アプリ")
st.write("車両番号や管理番号、会社名などで素早く検索できます。")

# データの読み込み＆会社名を下に引き継ぐ関数
@st.cache_data
def load_data():
    file_path = "新_登録車両資料_連動版.xlsx"
    df = pd.read_excel(file_path, sheet_name="新_登録車両資料", header=None)
    
    # 完全な空行を削除
    df = df.dropna(how="all")
    
    # 文字列に変換して空白を整理
    df = df.astype(str)
    df = df.replace(r'^\s*$', pd.NA, regex=True)
    
    # 会社名が入っている列（1列目）の値を、下の車両行に自動で引き継がせる（ffill）
    if len(df.columns) > 1:
        df[1] = df[1].ffill()
        
    return df

try:
    df = load_data()

    # 会社名（1列目）を一番左に持ってきて見やすくする
    if len(df.columns) > 1:
        cols = [1] + [col for col in df.columns if col != 1]
        df_display = df[cols]
    else:
        df_display = df

    # 検索ボックス
    search_query = st.text_input("🔍 キーワード検索", "")

    # フィルタリング処理
    filtered_df = df_display
    if search_query:
        mask = filtered_df.apply(lambda x: x.str.contains(search_query, case=False, na=False)).any(axis=1)
        filtered_df = filtered_df[mask]

    # ── 検索結果を数字の小さい順に自動で並び替える（ソート） ──
    try:
        # 左から3番目の列（車両番号などのデータが入っている列）を数値に変換して小さい順に並べる
        sort_target_col = filtered_df.columns[2] if len(filtered_df.columns) > 2 else filtered_df.columns[1]
        filtered_df = filtered_df.copy()
        filtered_df['_sort_val'] = pd.to_numeric(filtered_df[sort_target_col], errors='coerce')
        filtered_df = filtered_df.sort_values(by='_sort_val', ascending=True, na_position='last')
        filtered_df = filtered_df.drop(columns=['_sort_val'])
    except Exception:
        pass  # 万が一数値変換できなくてもそのまま表示する

    st.write(f"検索結果: **{len(filtered_df)}** 行（数字の小さい順に表示中）")

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
        if len(sheet_df.columns) > 1:
            sheet_df[1] = sheet_df[1].ffill()
            cols = [1] + [col for col in sheet_df.columns if col != 1]
            sheet_df = sheet_df[cols]
        st.dataframe(sheet_df, width="stretch", hide_index=True)

except Exception as e:
    st.error(f"データの読み込み中にエラーが発生しました: {e}")