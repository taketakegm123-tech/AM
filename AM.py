def dashboard(sheet1, sheet2):

    TYPE_JP = {
        "bank": "銀行",
        "cash": "現金",
        "invest": "投資"
    }

    TYPE_COLOR = {
        "bank": "#dbeafe",
        "cash": "#dcfce7",
        "invest": "#fde2e4"
    }

    st.markdown(
        """
        <style>
        .title-custom {
            font-size: 20px !important;
            font-weight: 600 !important;
            color: #555 !important;
            margin-bottom: 4px !important;
        }
        .big-card {
            padding: 10px 12px !important;
            border-radius: 12px;
            background-color: #e9d5ff !important;
            color: #555 !important;
            text-align: center;
            margin-bottom: 12px;
        }
        .big-card h1 {
            font-size: 20px !important;
            margin: 0 !important;
            line-height: 1.1 !important;
            color: #555 !important;
        }
        .big-card h2 {
            font-size: 14px !important;
            margin: 0 !important;
            line-height: 1.1 !important;
            color: #555 !important;
        }
        .cat-title {
            font-weight: bold;
            font-size: 18px;
            margin-bottom: 4px;
            color: #555 !important;
        }
        .cat-box {
            padding: 8px;
            border-radius: 8px;
            margin-bottom: 8px;
            border: 1px solid #ccc;
            color: #555 !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    # ★★★ タイトルをメニューより上に移動 ★★★
    st.markdown("<div class='title-custom'>資産管理</div>", unsafe_allow_html=True)

    current_total, history = calc_total_and_history(sheet1, sheet2)
    diff_day, diff_month = calc_diff(current_total, history)

    st.markdown(
        f"""
        <div class="big-card">
            <h1>総資産：¥{current_total:,.0f}</h1>
            <h2>前日比：{('+' if diff_day >= 0 else '')}¥{diff_day:,.0f} ｜ 前月比：{('+' if diff_month >= 0 else '')}¥{diff_month:,.0f}</h2>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.subheader("カテゴリ別資産")

    col_bank, col_cash, col_invest = st.columns(3)

    for col, t in zip([col_bank, col_cash, col_invest], ["bank", "cash", "invest"]):
        with col:
            st.markdown(f"<div class='cat-title'>{TYPE_JP[t]}</div>", unsafe_allow_html=True)

            df_cat = sheet1[sheet1["type"] == t]

            for _, row in df_cat.iterrows():
                st.markdown(
                    f"""
                    <div class="cat-box" style="background-color:{TYPE_COLOR[t]};">
                        {row['name']}<br>
                        ¥{row['balance']:,.0f}
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

    st.subheader("資産推移（概算）")
    st.line_chart(history.set_index("date")["total"])
