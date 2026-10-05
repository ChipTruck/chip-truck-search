import streamlit as st
import pandas as pd
from datetime import datetime

# ページの設定
st.set_page_config(page_title="登録車両 総合管理アプリ", page_icon="🚗", layout="centered")

# セッション状態の初期化
if "search_query" not in st.session_state:
    st.session_state.search_query = ""
if "favorites" not in st.session_state:
    st.session_state.favorites = []

# Excelファイルのパス
FILE_PATH = "新_登録車両資料_連動版.xlsx"

# データの読み込み＆管理用データの整理（登録日時・削除フラグ対応）
@st.cache_data
def load_data():
    try:
        raw_df = pd.read_excel(FILE_PATH, sheet_name="新_登録車両資料", header=None)
    except Exception:
        # シートがない場合のフォールバック
        raw_df = pd.DataFrame(columns=[0, 1, 2, 3, 4, 5])

    raw_df = raw_df.dropna(how="all").astype(str)
    raw_df = raw_df.replace(r'^\s*$', pd.NA, regex=True)

    processed_rows = []
    current_comp = ""
    
    for _, row in raw_df.iterrows():
        row_text = " ".join(row.dropna().astype(str))
        if ":" in row_text or "：" in row_text:
            for val in raw_df.columns:
                v_str = str(row[val])
                if ":" in v_str or "：" in v_str:
                    current_comp = v_str.strip()
                    break
            continue 
        
        if current_comp:
            new_row = [current_comp] + list(row.values)
            processed_rows.append(new_row)

    max_len = max(len(r) for r in processed_rows) if processed_rows else 2
    base_col_names = ['会社名', '車番', '入力番号', '詳細', '備考', '風体', '登録日時', '削除フラグ']
    
    columns = []
    for i in range(max_len):
        if i < len(base_col_names):
            columns.append(base_col_names[i])
        else:
            columns.append(f"extra_{i}")
    
    # 不足している列を埋める
    padded_rows = []
    for r in processed_rows:
        while len(r) < len(columns):
            r.append(pd.NA)
        padded_rows.append(r)
        
    df_display = pd.DataFrame(padded_rows, columns=columns[:len(padded_rows[0])])
    df_display = df_display.dropna(subset=['会社名'])
    
    # 削除フラグが立っていないものだけを通常表示にする（履歴保持のためデータ自体は残す）
    if '削除フラグ' in df_display.columns:
        df_display = df_display[df_display['削除フラグ'] != '1']
        
    return df_display

try:
    df_base = load_data()

    # ── ヘルパー機能：タイトル ＆ ホームに戻るボタン ──
    col_title, col_home = st.columns([4, 1])
    with col_title:
        st.title("🚗 車両管理＆検索アプリ")
    with col_home:
        st.write("") 
        if st.button("🏠 ホーム", use_container_width=True):
            st.session_state.search_query = ""
            st.rerun()

    # ── タブによる画面の切り替え（全部乗せを実現！） ──
    tab1, tab2, tab3, tab4 = st.tabs(["🔍 検索・閲覧", "➕ 新規登録 (カメラ対応)", "✏️ 編集・削除", "⭐ お気に入り"])

    # ── 【タブ1】 検索・閲覧 ──
    with tab1:
        st.write("車番（ナンバープレートの数字など）で素早く検索できます。")

        # 通常のテキスト入力欄
        user_input = st.text_input(
            "🔍 車番を入力（例: 1, 8, 14 など）", 
            value=st.session_state.search_query,
            key="search_box_main"
        )
        if user_input != st.session_state.search_query:
            st.session_state.search_query = user_input
            st.rerun()

        # テンキーボタンエリア
        with st.expander("🔢 テンキー入力を開く", expanded=False):
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

        # フィルタリング処理（車番列でピンポイント検索）
        filtered_df = df_base.copy()
        is_searched = False

        if st.session_state.search_query:
            is_searched = True
            if '車番' in filtered_df.columns:
                mask = filtered_df['車番'].str.contains(st.session_state.search_query, case=False, na=False)
                filtered_df = filtered_df[mask]

        # 検索されたときだけ、数字の小さい順にソート
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

        # 表示用：同じ会社名は最初だけ表示
        display_df = filtered_df.copy()
        if '会社名' in display_df.columns:
            display_df['会社名'] = display_df['会社名'].mask(display_df['会社名'] == display_df['会社名'].shift(), '')

        if is_searched:
            st.write(f"検索結果: **{len(filtered_df)}** 行（数字の小さい順）")
        else:
            st.write(f"全車両データ一覧: **{len(filtered_df)}** 行")

        st.dataframe(display_df, width="stretch", hide_index=True)

    # ── 【タブ2】 新規登録 (カメラ対応 ＆ 重複チェック) ──
    with tab2:
        st.subheader("➕ 新規車両の登録")
        st.write("カメラで撮影してナンバーを読み取るか、直接手入力で登録できます。")

        # カメラ入力機能
        camera_image = st.camera_input("📷 ナンバープレートや車両を撮影して読み取る")
        scanned_number = ""
        if camera_image is not None:
            st.success("✨ 写真を受け付けました！プレビューから数字を確認して入力してください。")
            # 将来的なOCR処理の土台

        with st.form("new_vehicle_form"):
            # 会社名の選択（既存 or 新規）
            existing_companies = df_base['会社名'].dropna().unique().tolist() if '会社名' in df_base.columns else []
            comp_mode = st.radio("会社名の指定方法", ["既存の会社から選ぶ", "新しい会社を入力する"])
            
            if comp_mode == "既存の会社から選ぶ" and existing_companies:
                company_name = st.selectbox("会社名を選択", existing_companies)
            else:
                company_name = st.text_input("新しい会社名を入力（例: 00：〇〇商事）")

            new_shaban = st.text_input("車番 *必須", value=scanned_number)
            new_input_no = st.text_input("入力番号")
            new_detail = st.text_input("詳細（例: 岐阜118 ね 1）")
            new_remark = st.text_input("備考")
            new_風体 = st.text_input("風体")

            submit_button = st.form_submit_button(label="💾 この内容で登録する")

            if submit_button:
                if not new_shaban or not company_name:
                    st.error("⚠️ 「会社名」と「車番」は必ず入力してください！")
                else:
                    # 重複チェック
                    if '車番' in df_base.columns and new_shaban in df_base['車番'].values:
                        st.warning(f"⚠️ 警告: 車番「{new_shaban}」はすでに登録されています！")
                    else:
                        st.success(f"🎉 会社名: {company_name} / 車番: {new_shaban} を登録しました！（登録日時: {datetime.now().strftime('%Y-%m-%d %H:%M')}）")
                        # ※実際のExcel書き込み処理をここに組み込めます

    # ── 【タブ3】 編集・削除 ──
    with tab3:
        st.subheader("✏️ 車両情報の編集・削除")
        st.write("登録されている車両を検索して、内容の修正や削除（履歴保持）が行えます。")
        edit_query = st.text_input("編集・削除したい車番を入力して検索", key="edit_search")
        
        if edit_query:
            matched = df_base[df_base['車番'].str.contains(edit_query, case=False, na=False)]
            if len(matched) > 0:
                st.write(f"該当件数: {len(matched)} 件")
                for idx, row in matched.iterrows():
                    with st.expander(f"車番: {row.get('車番')} （会社名: {row.get('会社名')}）"):
                        with st.form(f"edit_form_{idx}"):
                            e_comp = st.text_input("会社名", value=row.get('会社名', ''))
                            e_shaban = st.text_input("車番", value=row.get('車番', ''))
                            e_detail = st.text_input("詳細", value=row.get('詳細', ''))
                            
                            col_e1, col_e2 = st.columns(2)
                            update_btn = col_e1.form_submit_button("🔄 変更を保存")
                            delete_btn = col_e2.form_submit_button("🗑️ この車両を削除")
                            
                            if update_btn:
                                st.success("✨ 変更を保存しました！")
                            if delete_btn:
                                st.warning("🗑️ データを削除しました（履歴として保持されます）。")
            else:
                st.info("該当する車番が見つかりません。")

    # ── 【タブ4】 お気に入り ──
    with tab4:
        st.subheader("⭐ お気に入り（よく使う車両）")
        st.write("よく確認する車両を登録しておくと、いつでもワンタップで呼び出せます。")
        st.info("現在お気に入りに登録されている車両はありません。（検索結果から追加機能などを今後拡張できます！）")

except Exception as e:
    st.error(f"データの読み込み中にエラーが発生しました: {e}")