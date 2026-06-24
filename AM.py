import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

EXCEL_PATH = "assets.xlsx"

# ---------- データ読み込み ----------
@st.cache_data
def load_data():
    sheet1 = pd.read_excel(EXCEL_PATH, sheet_name="Sheet1")
    sheet2 = pd.read_excel(EXCEL_PATH, sheet_name="Sheet2")
    # 列名を統一
    sheet1.columns = ["type", "name", "balance"]
    sheet2.columns = ["date", "from", "to", "amount", "memo"]
    sheet2["date"] = pd.to_datetime(sheet2["date"])
    return sheet1, sheet2

# ---------- 総資産と日次推移の計算 ----------
def calc_total_and_history(sheet1, sheet2):
    # 現在の総資産
    current_total = sheet1["balance"].sum()

    # 資産カテゴリ一覧（Sheet1 の type）
    asset_types = sheet1["type"].unique().tolist()

    # 日次の総資産変化量（支出・収入のみ）
    # 振替（from と to 両方が資産）は総資産変化 0
    def row_delta(row):
        f = row["from"]
        t = row["to"]
        amt = row["amount"]
        f_is_asset = f in asset_types
        t_is_asset = t in asset_types

        # 振替：総資産変化なし
        if f_is_asset and t_is_asset:
            return 0
        # 支出：資産 → 費目
        if f_is_asset and not t_is_asset:
            return -amt
        # 収入：収入源 → 資産
        if not f_is_asset and t_is_asset:
            return +amt
        # それ以外は総資産に影響なし
        return 0

    sheet2["delta"] = sheet2.apply(row_delta, axis=1)

    # 日付ごとの変化量
    daily_delta = sheet2.groupby("date")["delta"].sum().sort_index()

    # ここでは「現在の総資産」を基準に、
    # 過去の総資産を「逆算」する簡易版を使う
    # （厳密な過去残高再現は、初期残高の定義が必要になるため）
    # 今日を max(date) とし、その日までの変化を current_total に対応させる
    if len(daily_delta) == 0:
        # 履歴がない場合
        history = pd.DataFrame({
            "date": [datetime.today().date()],
            "total": [current_total]
        })
    else:
        last_date = daily_delta.index.max().date()
        # last_date 時点の総資産 = current_total とみなす
        # そこから過去に向かって累積を逆算
        daily_delta_sorted = daily_delta.sort_index(ascending=False)
        totals = []
        running_total = current_total
        for d, delta in daily_delta_sorted.items():
            totals.append((d.date(), running_total))
            running_total -= delta  # 1日前に戻るイメージ
        history = pd.DataFrame(totals, columns=["date", "total"]).sort_values("date")

    return current_total, history

# ---------- 前日比・前月比の計算 ----------
def calc_diff(current_total, history):
    if history.empty:
        return 0, 0

    today = history["date"].max()
    # 前日
    yesterday = today - timedelta(days=1)
    # 先月末（ざっくり：今日の1ヶ月前に最も近い日）
    last_month = today - timedelta(days=30)

    # 前日総資産
    y_row = history[history["date"] <= yesterday]
    if len(y_row) > 0:
        yesterday_total = y_row["total"].iloc[-1]
    else:
        yesterday_total = current_total

    # 先月総資産
    m_row = history[history["date"] <= last_month]
    if len(m_row) > 0:
        last_month_total = m_row["total"].iloc[-1]
    else:
        last_month_total = current_total

    diff_day = current_total - yesterday_total
    diff_month = current_total - last_month_total
    return diff_day, diff_month

# ---------- スマホ UI 用の Dashboard ----------
def dashboard(sheet1, sheet2):
    st.markdown(
        """
        <style>
        .big-card {
            padding: 16px;
            border-radius: 12px;
            background-color: #0f172a;
            color: white;
            box-shadow: 0 4px 12px rgba(0,0,0,0.3);
            text-align: center;
            margin-bottom: 16px;
        }
        .big-card h1 {
            font-size: 28px;
            margin: 0;
        }
        .big-card h2 {
            font-size: 18px;
            margin: 4px 0;
        }
        .small-card {
            padding: 12px;
            border-radius: 10px;
            background-color: #1e293b;
            color: white;
            margin-right: 8px;
            min-width: 140px;
        }
        .small-card h3 {
            font-size: 16px;
            margin: 0 0 4px 0;
        }
        .small-card p {
            font-size: 14px;
            margin: 0;
        }
        .scroll-row {
            display: flex;
            overflow-x: auto;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    sheet1, sheet2 = sheet1.copy(), sheet2.copy()
    current_total, history = calc_total_and_history(sheet1, sheet2)
    diff_day, diff_month = calc_diff(current_total, history)

    # 総資産カード
    st.markdown(
        f"""
        <div class="big-card">
            <h1>総資産：¥{current_total:,.0f}</h1>
            <h2>前日比：{('+' if diff_day >= 0 else '')}¥{diff_day:,.0f}</h2>
            <h2>前月比：{('+' if diff_month >= 0 else '')}¥{diff_month:,.0f}</h2>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # カテゴリ別サマリー
    st.markdown("#### カテゴリ別資産")
    cat_summary = sheet1.groupby("type")["balance"].sum().reset_index()

    cards_html = '<div class="scroll-row">'
    for _, row in cat_summary.iterrows():
        cards_html += f"""
        <div class="small-card">
            <h3>{row['type']}</h3>
            <p>¥{row['balance']:,.0f}</p>
        </div>
        """
    cards_html += "</div>"

    st.markdown(cards_html, unsafe_allow_html=True)

    # 資産推移グラフ
    st.markdown("#### 資産推移（概算）")
    st.line_chart(history.set_index("date")["total"])

# ---------- メイン ----------
def main():
    st.set_page_config(page_title="資産管理アプリ", layout="centered")
    sheet1, sheet2 = load_data()

    # 下部タブ風メニュー（実際はラジオボタンで切り替え）
    menu = st.radio(
        "メニュー",
        ["🏠 Dashboard", "➕ Input", "📄 List", "📊 Charts"],
        horizontal=True,
    )

    if menu == "🏠 Dashboard":
        dashboard(sheet1, sheet2)
    elif menu == "➕ Input":
        st.write("入力画面（ここは後でスマホ最適化版を作る）")
    elif menu == "📄 List":
        st.write("一覧画面（ここも後でカード型にできる）")
    elif menu == "📊 Charts":
        st.write("グラフ画面（資産推移やカテゴリ別円グラフを追加予定）")

if __name__ == "__main__":
    main()
