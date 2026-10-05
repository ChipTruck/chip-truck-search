# ── 【タブ1】 検索・閲覧 ──
    with tab1:
        st.write("車番の数字で素早く検索できます。")

        # セッション状態のキーを統一
        if "search_box_main" not in st.session_state:
            st.session_state.search_box_main = ""

        # 入力ボックス
        user_input = st.text_input(
            "🔍 車番を入力（例: 1, 8, 14 など）", 
            key="search_box_main"
        )

        # テンキー操作（1〜9, 0, 🎤マイク, クリア, 決定）
        with st.expander("🔢 テンキー入力を開く", expanded=True):
            r1c1, r1c2, r1c3 = st.columns(3)
            if r1c1.button("1", use_container_width=True): 
                st.session_state.search_box_main += "1"
                st.rerun()
            if r1c2.button("2", use_container_width=True): 
                st.session_state.search_box_main += "2"
                st.rerun()
            if r1c3.button("3", use_container_width=True): 
                st.session_state.search_box_main += "3"
                st.rerun()

            r2c1, r2c2, r2c3 = st.columns(3)
            if r2c1.button("4", use_container_width=True): 
                st.session_state.search_box_main += "4"
                st.rerun()
            if r2c2.button("5", use_container_width=True): 
                st.session_state.search_box_main += "5"
                st.rerun()
            if r2c3.button("6", use_container_width=True): 
                st.session_state.search_box_main += "6"
                st.rerun()

            r3c1, r3c2, r3c3 = st.columns(3)
            if r3c1.button("7", use_container_width=True): 
                st.session_state.search_box_main += "7"
                st.rerun()
            if r3c2.button("8", use_container_width=True): 
                st.session_state.search_box_main += "8"
                st.rerun()
            if r3c3.button("9", use_container_width=True): 
                st.session_state.search_box_main += "9"
                st.rerun()

            r4c1, r4c2, r4c3 = st.columns(3)
            if r4c1.button("0", use_container_width=True): 
                st.session_state.search_box_main += "0"
                st.rerun()
            if r4c2.button("🎤 音声入力", use_container_width=True): 
                st.info("💡 スマホのキーボードのマイクを使うか、キーボードから音声入力できます")
            if r4c3.button("クリア", use_container_width=True): 
                st.session_state.search_box_main = ""
                st.rerun()

            # 番号決定（検索実行）ボタン
            if st.button("🔍 番号決定（検索）", use_container_width=True, type="primary"):
                st.rerun()

        search_val = st.session_state.search_box_main
        filtered_df = df_base.copy()
        is_searched = bool(search_val)

        if is_searched:
            if '車番' in filtered_df.columns:
                mask = filtered_df['車番'].str.contains(search_val, case=False, na=False)
                filtered_df = filtered_df[mask]

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
            st.write(f"検索結果: **{len(filtered_df)}** 行 （検索キー: {search_val}）")
        else:
            st.write(f"全車両一覧: **{len(filtered_df)}** 行")

        st.dataframe(display_df, use_container_width=True, hide_index=True)