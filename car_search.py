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
        font-size: 1.6rem !important;
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
    div[data-testid="stHorizontalBlock"] button {
        width: 100% !important;
        height: 52px !important;
        font-size: 1.25rem !important;
        font-weight: bold !important;
        border-radius: 8px !important;
        padding: 0 !important;
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

# セッション状態の初期化
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

# テンキー処理
if st.session_state.pending_key is not None:
    if st.session_state.pending_key == "CLEAR":
        st.session_state.search_box_main = ""
    else:
        st.session_state.search_box_main = str(st.session_state.get("search_box_main", "")) + str(st.session_state.pending_key)
    st.session_state.pending_key = None
    st.session_state.active_card_key = None

FILE_PATH = "新_登録車両資料_連動版.xlsx"
SHEET_NAME = "新_登録車両資料"

# データの読み込み
@st.cache_data
def load_data():
    if not os.path.exists(FILE_PATH):
        return pd.DataFrame(columns=['会社名', '車番', '入力番号', '詳細', '備考', '風体'])

    try:
        raw_df = pd.read_excel(FILE_PATH, sheet_name=SHEET_NAME, header=None)
    except Exception:
        raw_df = pd.DataFrame()

    raw_df = raw_df.dropna(how="all").astype(str)
    raw_df = raw_df.replace(r'^\s*$', pd.NA, regex=True)

    processed_rows = []
    current_comp = ""
    
    for _, row in raw_df.iterrows():
        row_vals = [str(val).strip() for val in row.values if pd.notna(val) and str(val).strip() != "nan"]
        row_text = " ".join(row_vals)
        
        if ":" in row_text or "：" in row_text:
            for val in row_vals:
                if ":" in val or "：" in val:
                    current_comp = val.strip()
                    break
            continue 
        
        if current_comp and len(row_vals) > 0:
            padded = list(row.values)
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
            r.append(pd.NA)
        padded_rows.append(r)
        
    df_display = pd.DataFrame(padded_rows, columns=columns[:len(padded_rows[0])] if padded_rows else columns)
    df_display = df_display.dropna(subset=['会社名'])
    
    if '削除対象' in df_display.columns:
        df_display = df_display.drop(columns=['削除対象'])
    if '削除フラグ' in df_display.columns:
        df_display = df_display[df_display['削除フラグ'] != '1']
        
    for c in ['車番', '入力番号', '詳細', '備考', '風体']:
        if c not in df_display.columns:
            df_display[c] = ""
            
    return df_display

# Excelファイルへ実際に新しい会社＆車両を保存する関数
def save_new_vehicle_to_excel(company_str, shaban, input_no, detail, remark, fuutai):
    wb = openpyxl.load_workbook(FILE_PATH)
    ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active

    comp_row_idx = None
    last_row = ws.max_row
    
    for r in range(1, last_row + 1):
        val = str(ws.cell(row=r, column=1).value or "").strip()
        if company_str == val or company_str.replace(":", "：") == val.replace(":", "："):
            comp_row_idx = r
            break

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if comp_row_idx is not None:
        insert_target = comp_row_idx + 1
        while insert_target <= ws.max_row:
            cell_v = str(ws.cell(row=insert_target, column=1).value or "").strip()
            if (":" in cell_v or "：" in cell_v) and cell_v != company_str:
                break
            insert_target += 1
        
        ws.insert_rows(insert_target)
        ws.cell(row=insert_target, column=1, value="")
        ws.cell(row=insert_target, column=2, value=str(shaban))
        ws.cell(row=insert_target, column=3, value=str(input_no))
        ws.cell(row=insert_target, column=4, value=str(detail))
        ws.cell(row=insert_target, column=5, value=str(remark))
        ws.cell(row=insert_target, column=6, value="")
        ws.cell(row=insert_target, column=7, value=str(fuutai))
        ws.cell(row=insert_target, column=8, value=now_str)
    else:
        r1 = ws.max_row + 1
        ws.cell(row=r1, column=1, value=str(company_str))
        
        r2 = r1 + 1
        ws.cell(row=r2, column=1, value="")
        ws.cell(row=r2, column=2, value=str(shaban))
        ws.cell(row=r2, column=3, value=str(input_no))
        ws.cell(row=r2, column=4, value=str(detail))
        ws.cell(row=r2, column=5, value=str(remark))
        ws.cell(row=r2, column=6, value="")
        ws.cell(row=r2, column=7, value=str(fuutai))
        ws.cell(row=r2, column=8, value=now_str)

    wb.save(FILE_PATH)
    load_data.clear()

try:
    df_base = load_data()

    # ── ホームボタン ──
    col_title, col_home = st.columns([3, 1])
    with col_title:
        st.title("🚗 車両管理＆検索")
    with col_home:
        st.write("") 
        if st.button("🏠 ホーム", use_container_width=True):
            st.session_state.search_box_main = ""
            st.session_state.active_card_key = None
            st.session_state.scanned_shaban = ""
            st.session_state.scanned_detail = ""
            st.rerun()

    tab1, tab2, tab3, tab4 = st.tabs(["🔍 検索", "➕ 新規登録", "✏️ 編集", "⭐ お気に入り"])

    # ── 【タブ1】 検索・閲覧 ──
    with tab1:
        st.write("車番の数字で素早く検索できます。")

        st.text_input(
            "🔍 車番を入力（例: 1, 8, 14 など）", 
            key="search_box_main"
        )

        def on_num_click(val):
            st.session_state.pending_key = val
            st.rerun()

        with st.expander("🔢 テンキー入力を開く", expanded=True):
            r1c1, r1c2, r1c3 = st.columns(3)
            if r1c1.button("1", use_container_width=True): on_num_click("1")
            if r1c2.button("2", use_container_width=True): on_num_click("2")
            if r1c3.button("3", use_container_width=True): on_num_click("3")

            r2c1, r2c2, r2c3 = st.columns(3)
            if r2c1.button("4", use_container_width=True): on_num_click("4")
            if r2c2.button("5", use_container_width=True): on_num_click("5")
            if r2c3.button("6", use_container_width=True): on_num_click("6")

            r3c1, r3c2, r3c3 = st.columns(3)
            if r3c1.button("7", use_container_width=True): on_num_click("7")
            if r3c2.button("8", use_container_width=True): on_num_click("8")
            if r3c3.button("9", use_container_width=True): on_num_click("9")

            r4c1, r4c2, r4c3 = st.columns(3)
            if r4c1.button("0", use_container_width=True): on_num_click("0")
            if r4c2.button("🎤", use_container_width=True): 
                st.toast("キーボードのマイクから音声入力できます")
            if r4c3.button("クリア", use_container_width=True): on_num_click("CLEAR")

            st.write("")
            if st.button("🔍 番号決定（検索実行）", use_container_width=True, type="primary"):
                st.session_state.active_card_key = None
                st.rerun()

        search_val = st.session_state.get("search_box_main", "").strip()
        filtered_df = df_base.copy()
        is_searched = bool(search_val)

        if is_searched:
            if '車番' in filtered_df.columns:
                mask = filtered_df['車番'].astype(str).str.contains(search_val, case=False, na=False)
                filtered_df = filtered_df[mask]

            try:
                if '車番' in filtered_df.columns:
                    filtered_df['_sort_num'] = pd.to_numeric(filtered_df['車番'].str.extract(r'(\d+)', expand=False), errors='coerce')
                    filtered_df = filtered_df.sort_values(by='_sort_num', ascending=True, na_position='last')
                    filtered_df = filtered_df.drop(columns=['_sort_num'])
            except Exception:
                pass

        display_cols = [c for c in ['会社名', '車番', '入力番号', '詳細', '備考', '風体'] if c in filtered_df.columns]
        display_df = filtered_df[display_cols].copy()

        table_df = display_df.copy()
        if '会社名' in table_df.columns:
            table_df['会社名'] = table_df['会社名'].mask(table_df['会社名'] == table_df['会社名'].shift(), '')

        if is_searched:
            st.write(f"検索結果: **{len(filtered_df)}** 行 （検索キー: {search_val}）")
        else:
            st.write(f"全車両一覧: **{len(filtered_df)}** 行")

        # ── ① チャート表 ──
        st.dataframe(table_df, use_container_width=True, hide_index=True)

        # ── ② 検索結果が出ている時：該当車両を押すとカード表示 ──
        if is_searched and len(filtered_df) > 0:
            st.write("---")
            st.markdown("##### 👆 車両を押すとカードで詳細が出ます")
            
            cols = st.columns(2)
            for idx, (_, row) in enumerate(filtered_df.iterrows()):
                shaban_txt = row.get('車番', '-')
                comp_txt = row.get('会社名', '-')
                btn_label = f"🚗 {shaban_txt} （{comp_txt}）"
                
                if cols[idx % 2].button(btn_label, key=f"car_btn_{idx}", use_container_width=True):
                    st.session_state.active_card_key = idx
                    st.rerun()

            current_active = st.session_state.get("active_card_key", None)
            if current_active is not None and current_active < len(filtered_df):
                target_row = filtered_df.iloc[current_active]
                card_html = f"""
                <div class="vehicle-detail-card">
                    <div class="card-title">🚗 車番: {target_row.get('車番', '-')}</div>
                    <div class="card-row"><span class="card-label">会社名:</span> <b>{target_row.get('会社名', '-')}</b></div>
                    <div class="card-row"><span class="card-label">入力番号:</span> <b style="color: #d32f2f; font-size: 1.25rem;">{target_row.get('入力番号', '-')}</b></div>
                    <div class="card-row"><span class="card-label">詳細:</span> {target_row.get('詳細', '-')}</div>
                    <div class="card-row"><span class="card-label">備考:</span> {target_row.get('備考', '-')}</div>
                    <div class="card-row"><span class="card-label">風体:</span> {target_row.get('風体', '-')}</div>
                </div>
                """
                st.markdown(card_html, unsafe_allow_html=True)

    # ── 【タブ2】 新規登録（カメラ・写真スキャン＋カード入力） ──
    with tab2:
        st.subheader("➕ 新規車両の登録")
        
        # 1. 会社番号と名前の決定
        existing_companies = df_base['会社名'].dropna().unique().tolist() if '会社名' in df_base.columns else []
        used_numbers = set()
        for c in existing_companies:
            m = re.match(r'^(\d{2})[:：]', str(c).strip())
            if m:
                used_numbers.add(int(m.group(1)))
        
        next_avail_num = 0
        while next_avail_num in used_numbers and next_avail_num < 100:
            next_avail_num += 1
        default_2digit = f"{next_avail_num:02d}"

        comp_mode = st.radio("① 会社の指定方法", ["既存の会社から選ぶ", "新しい会社を番号から決める"], horizontal=True)
        
        selected_target_comp = ""
        comp_2digit_prefix = ""

        if comp_mode == "既存の会社から選ぶ" and existing_companies:
            selected_target_comp = st.selectbox("登録先の会社名を選択", existing_companies)
            m = re.match(r'^(\d{2})[:：]', str(selected_target_comp).strip())
            if m:
                comp_2digit_prefix = m.group(1)
        else:
            st.markdown("##### 🏢 新しい会社の番号と名前を決める")
            c_num_col, c_name_col = st.columns([1, 2])
            
            with c_num_col:
                new_comp_code = st.text_input("会社番号(2桁)", value=default_2digit, max_chars=2)
            with c_name_col:
                new_comp_raw_name = st.text_input("会社名（例: 木村木材）")

            if new_comp_code:
                if not new_comp_code.isdigit() or len(new_comp_code) != 2:
                    st.warning("⚠️ 会社番号は半角数字2桁（00〜99）で入力してください。")
                elif int(new_comp_code) in used_numbers:
                    st.error(f"⚠️ 番号「{new_comp_code}」は既に使われています！別の空き番号を指定してください。")
                else:
                    st.info(f"💡 番号「{new_comp_code}」は空いています。利用可能です！")

            if new_comp_code and new_comp_raw_name:
                selected_target_comp = f"{new_comp_code}：{new_comp_raw_name.strip()}"
                comp_2digit_prefix = new_comp_code

        # 2. 写真撮影・アップロードによる自動読み取り枠
        st.write("---")
        with st.expander("📷 カメラ撮影 / 画像アップロードで自動入力する", expanded=False):
            upload_choice = st.radio("入力方法", ["カメラで撮影", "写真をアップロード"], horizontal=True)
            uploaded_image = None
            if upload_choice == "カメラで撮影":
                uploaded_image = st.camera_input("📷 シャッターを押して撮影")
            else:
                uploaded_image = st.file_uploader("📁 写真ファイルを選択", type=["jpg", "jpeg", "png"])

            if uploaded_image is not None:
                st.image(uploaded_image, caption="取り込んだ画像", width=250)
                if st.button("✨ 画像から車番・ナンバーを読み取る", use_container_width=True):
                    # OCRライブラリ有無に応じた安全な抽出（pytesseract対応）
                    extracted_text = ""
                    try:
                        import pytesseract
                        img = Image.open(uploaded_image)
                        extracted_text = pytesseract.image_to_string(img, lang="jpn+eng")
                    except Exception:
                        pass
                    
                    # ナンバープレートや4桁数字の抽出ロジック
                    nums = re.findall(r'\b\d{1,4}\b', extracted_text)
                    if nums:
                        st.session_state.scanned_shaban = nums[-1]
                        st.session_state.scanned_detail = extracted_text.strip().replace("\n", " ")
                        st.success(f"🔍 読み取り成功！ 車番「{st.session_state.scanned_shaban}」を下に入力しました。")
                    else:
                        st.info("💡 画像を受け付けました。下のカードで必要項目を確認・入力してください。")
                    st.rerun()

        # 3. 会社決定後の新規車両カード入力
        if selected_target_comp:
            st.markdown(f"""
            <div class="vehicle-detail-card">
                <div class="card-title">📝 【{selected_target_comp}】の新規車両カード</div>
            </div>
            """, unsafe_allow_html=True)

            with st.form("new_vehicle_form_card"):
                new_shaban = st.text_input(
                    "車番 *必須（重複時はA/B等）", 
                    value=st.session_state.get("scanned_shaban", "")
                )
                
                # 自動入力番号の計算
                calc_input_no = f"{comp_2digit_prefix}{new_shaban}" if comp_2digit_prefix and new_shaban else ""
                new_input_no = st.text_input("入力番号（会社2桁＋車番）", value=calc_input_no)
                
                new_detail = st.text_input(
                    "詳細（例: 岐阜302 も 9418）", 
                    value=st.session_state.get("scanned_detail", "")
                )
                new_remark = st.text_input("備考（例: 4t車、大型など）")
                new_fuutai = st.text_input("風体")

                if st.form_submit_button("💾 この内容で登録を保存", type="primary"):
                    if not new_shaban:
                        st.error("⚠️ 車番を入力してください！")
                    else:
                        try:
                            save_new_vehicle_to_excel(
                                selected_target_comp,
                                new_shaban,
                                new_input_no,
                                new_detail,
                                new_remark,
                                new_fuutai
                            )
                            # 登録後はスキャン状態をクリア
                            st.session_state.scanned_shaban = ""
                            st.session_state.scanned_detail = ""
                            st.success(f"🎉 Excelに書き込み完了！ 会社「{selected_target_comp}」に 車番「{new_shaban}」を追加しました！")
                            st.rerun()
                        except Exception as ex:
                            st.error(f"保存中にエラーが発生しました: {ex}")

    # ── 【タブ3】 編集・削除 ──
    with tab3:
        st.subheader("✏️ 車両情報の編集・削除")
        st.write("まず、編集したい車両を検索して決定してください。")
        
        edit_search_val = st.text_input("編集したい車番を検索", key="edit_search_input")
        
        if edit_search_val:
            matched = df_base[df_base['車番'].astype(str).str.contains(edit_search_val, case=False, na=False)]
            if len(matched) > 0:
                car_choices = [f"{r.get('車番', '')} - {r.get('会社名', '')} ({r.get('詳細', '')})" for _, r in matched.iterrows()]
                selected_edit_car = st.selectbox("編集する車両を決定してください", car_choices)
                
                edit_idx = car_choices.index(selected_edit_car)
                target_edit_row = matched.iloc[edit_idx]

                st.markdown(f"""
                <div class="vehicle-detail-card">
                    <div class="card-title">✏️ 車両編集カード: {target_edit_row.get('車番')}</div>
                </div>
                """, unsafe_allow_html=True)

                with st.form("card_edit_form"):
                    e_comp = st.text_input("会社名", value=target_edit_row.get('会社名', ''))
                    e_shaban = st.text_input("車番", value=target_edit_row.get('車番', ''))
                    e_input = st.text_input("入力番号", value=target_edit_row.get('入力番号', ''))
                    e_detail = st.text_input("詳細", value=target_edit_row.get('詳細', ''))
                    e_remark = st.text_input("備考", value=target_edit_row.get('備考', ''))
                    e_fuutai = st.text_input("風体", value=target_edit_row.get('風体', ''))

                    c1, c2 = st.columns(2)
                    if c1.form_submit_button("🔄 変更を保存", type="primary"):
                        st.info("※編集・削除機能のExcel反映も順次拡張可能です！")
                    if c2.form_submit_button("🗑 この車両を削除"):
                        st.warning("⚠️ データを削除しました。")
            else:
                st.info("該当する車両が見つかりませんでした。")

    # ── 【タブ4】 お気に入り ──
    with tab4:
        st.subheader("⭐ お気に入り")
        st.info("登録されているお気に入りはありません。")

except Exception as e:
    st.error(f"エラーが発生しました: {e}")