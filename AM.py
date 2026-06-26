import streamlit as st
import msal
import requests
import pandas as pd
import io
from datetime import datetime, timedelta

# ============================
# 1. Secrets 読み込み
# ============================
CLIENT_ID = st.secrets["azure"]["client_id"]
CLIENT_SECRET = st.secrets["azure"]["client_secret"]
TENANT_ID = st.secrets["azure"]["tenant_id"]
REDIRECT_URI = st.secrets["onedrive"]["redirect_uri"]

FILE_PATH = "Asset_Manager/assets.xlsx"
SCOPE = ["Files.ReadWrite"]

# ============================
# 2. MSAL 関連
# ============================
def load_cache():
    if "token_cache" not in st.session_state:
        st.session_state["token_cache"] = msal.SerializableTokenCache()
    return st.session_state["token_cache"]

def save_cache(cache):
    st.session_state["token_cache"] = cache

def build_msal_app(cache=None):
    return msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority="https://login.microsoftonline.com/consumers",
        client_credential=CLIENT_SECRET,
        token_cache=cache,
    )

def get_token(show_login_ui=True):
    cache = load_cache()
    app = build_msal_app(cache)

    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPE, account=accounts[0])
        if result and "access_token" in result:
            return result["access_token"]

    raw_params = st.query_params
    query_params = dict(raw_params)

    if "code" in query_params:
        code = query_params["code"]
        result = app.acquire_token_by_authorization_code(
            code,
            scopes=SCOPE,
            redirect_uri=REDIRECT_URI,
        )
        if "access_token" in result:
            save_cache(cache)
            return result["access_token"]

    if show_login_ui:
        return app.get_authorization_request_url(SCOPE, redirect_uri=REDIRECT_URI)

    return None

# ============================
# 3. OneDrive 読み込み／書き込み
# ============================
def read_workbook_from_onedrive(access_token, file_path):
    url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{file_path}:/content"
    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        bio = io.BytesIO(response.content)
        sheets = pd.read_excel(bio, sheet_name=None)
        return sheets
    else:
        st.error("ファイル取得に失敗しました")
        st.write(response.text)
        return None

def write_workbook_to_onedrive(access_token, file_path, sheets_dict):
    url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{file_path}:/content"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in sheets_dict.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    output.seek(0)

    response = requests.put(url, headers=headers, data=output.read())
    return response.status_code in [200, 201]

# ============================
# 4. Sheet2 の列を統一
# ============================
def clean_sheet2(df):
    expected = ["date", "type", "amount", "from", "to", "memo"]

    while len(df.columns) < len(expected):
        df[expected[len(df.columns)]] = ""

    if len(df.columns) > len(expected):
        df = df.iloc[:, :len(expected)]

    df.columns = expected
    return df

# ============================
# 5. Dashboard 計算ロジック
# ============================
def calc_income_expense(sheet2):
    df = sheet2.copy()
    df["date"] = pd.to_datetime(df["date"])

    # 振替は収支に含めない
    df = df[df["type"] != "振替"]

    today = datetime.today().date()
    yesterday = today - timedelta(days=1)
    month_start = today.replace(day=1)

    df_y = df[df["date"].dt.date == yesterday]
    yesterday_total = df_y["amount"].sum()

    df_m = df[df["date"].dt.date >= month_start]
    month_total = df_m["amount"].sum()

    return yesterday_total, month_total

def calc_total(sheet1):
    sheet1 = sheet1.copy()
    sheet1.columns = ["type", "name", "balance"]
    return sheet1["balance"].sum()

# ============================
# 6. Dashboard ページ
# ============================
def dashboard_page(sheet1, sheet2):
    total = calc_total(sheet1)
    y_total, m_total = calc_income_expense(sheet2)

    st.markdown(
        """
        <style>
        .card {
            padding: 12px;
            border-radius: 10px;
            background-color: #e9d5ff;
            border: 1px solid #aaa;
            color: #555;
            text-align: center;
            margin-bottom: 12px;
        }
        .card h1 { font-size: 20px; margin: 0; }
        .card h2 { font-size: 14px; margin: 0; }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(f"<div class='card'><h1>総資産：¥{total:,.0f}</h1></div>", unsafe_allow_html=True)
    st.markdown(f"<div class='card'><h2>昨日の収支：¥{y_total:,.0f}</h2></div>", unsafe_allow_html=True)
    st.markdown(f"<div class='card'><h2>今月の収支：¥{m_total:,.0f}</h2></div>", unsafe_allow_html=True)

    TYPE_COLOR = {"bank": "#dbeafe", "cash": "#dcfce7", "invest": "#fde2e4"}
    col1, col2, col3 = st.columns(3)

    for col, t in zip([col1, col2, col3], ["bank", "cash", "invest"]):
        df_cat = sheet1[sheet1["type"] == t]
        for _, row in df_cat.iterrows():
            col.markdown(
                f"""
                <div style="padding:10px; border-radius:8px; border:1px solid #aaa;
                background-color:{TYPE_COLOR[t]}; margin-bottom:10px;">
                    {row['name']}<br>¥{row['balance']:,.0f}
                </div>
                """,
                unsafe_allow_html=True,
            )

# ============================
# 7. Input ページ（支出 / 収入 / 振替）
# ============================
def input_page(sheet1, sheet2, sheets, token):

    if "exp_reset" not in st.session_state:
        st.session_state.exp_reset = 0
    if "inc_reset" not in st.session_state:
        st.session_state.inc_reset = 0
    if "trf_reset" not in st.session_state:
        st.session_state.trf_reset = 0

    asset_list = sheet1["name"].tolist()
    default_wallet_index = asset_list.index("財布") if "財布" in asset_list else 0

    # ============================
    # 支出
    # ============================
    with st.container(border=True):
        st.subheader("支出")

        k = st.session_state.exp_reset

        exp_date = st.date_input("日付", key=f"exp_date_{k}")
        exp_amount = st.text_input("金額", key=f"exp_amount_{k}")
        exp_from = st.selectbox("from（出金元）", asset_list, index=default_wallet_index, key=f"exp_from_{k}")
        exp_to = st.text_input("to（費目）", key=f"exp_to_{k}")
        exp_memo = st.text_input("メモ", key=f"exp_memo_{k}")

        if st.button("支出を登録"):
            if not exp_amount.isdigit():
                st.error("金額は数字で入力してください")
                return

            amount_val = -abs(int(exp_amount))

            new_row = {
                "date": pd.to_datetime(exp_date),
                "type": "支出",
                "amount": amount_val,
                "from": exp_from,
                "to": exp_to,
                "memo": exp_memo
            }

            df = sheet2.copy()
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            sheets["Sheet2"] = df

            s1 = sheet1.copy()
            s1.columns = ["type", "name", "balance"]
            s1.loc[s1["name"] == exp_from, "balance"] -= abs(int(exp_amount))
            sheets["Sheet1"] = s1

            write_workbook_to_onedrive(token, FILE_PATH, sheets)
            st.session_state.sheets = read_workbook_from_onedrive(token, FILE_PATH)

            st.session_state.exp_reset += 1
            st.rerun()

    # ============================
    # 収入
    # ============================
    with st.container(border=True):
        st.subheader("収入")

        k = st.session_state.inc_reset

        inc_date = st.date_input("日付", key=f"inc_date_{k}")
        inc_amount = st.text_input("金額", key=f"inc_amount_{k}")
        inc_from = st.text_input("from（収入元）", key=f"inc_from_{k}")
        inc_to = st.selectbox("to（入金先資産）", asset_list, index=default_wallet_index, key=f"inc_to_{k}")
        inc_memo = st.text_input("メモ", key=f"inc_memo_{k}")

        if st.button("収入を登録"):
            if not inc_amount.isdigit():
                st.error("金額は数字で入力してください")
                return

            amount_val = abs(int(inc_amount))

            new_row = {
                "date": pd.to_datetime(inc_date),
                "type": "収入",
                "amount": amount_val,
                "from": inc_from,
                "to": inc_to,
                "memo": inc_memo
            }

            df = sheet2.copy()
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            sheets["Sheet2"] = df

            s1 = sheet1.copy()
            s1.columns = ["type", "name", "balance"]
            s1.loc[s1["name"] == inc_to, "balance"] += abs(int(inc_amount))
            sheets["Sheet1"] = s1

            write_workbook_to_onedrive(token, FILE_PATH, sheets)
            st.session_state.sheets = read_workbook_from_onedrive(token, FILE_PATH)

            st.session_state.inc_reset += 1
            st.rerun()

    # ============================
    # 振替（Transfer）
    # ============================
    with st.container(border=True):
        st.subheader("振替（資産移動）")

        k = st.session_state.trf_reset

        trf_date = st.date_input("日付", key=f"trf_date_{k}")
        trf_amount = st.text_input("金額", key=f"trf_amount_{k}")
        trf_from = st.selectbox("from（出金元）", asset_list, key=f"trf_from_{k}")
        trf_to = st.selectbox("to（入金先）", asset_list, key=f"trf_to_{k}")
        trf_memo = st.text_input("メモ", key=f"trf_memo_{k}")

        if st.button("振替を登録"):
            if not trf_amount.isdigit():
                st.error("金額は数字で入力してください")
                return

            amount_val = abs(int(trf_amount))

            new_row = {
                "date": pd.to_datetime(trf_date),
                "type": "振替",
                "amount": amount_val,
                "from": trf_from,
                "to": trf_to,
                "memo": trf_memo
            }

            df = sheet2.copy()
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            sheets["Sheet2"] = df

            s1 = sheet1.copy()
            s1.columns = ["type", "name", "balance"]
            s1.loc[s1["name"] == trf_from, "balance"] -= amount_val
            s1.loc[s1["name"] == trf_to, "balance"] += amount_val
            sheets["Sheet1"] = s1

            write_workbook_to_onedrive(token, FILE_PATH, sheets)
            st.session_state.sheets = read_workbook_from_onedrive(token, FILE_PATH)

            st.session_state.trf_reset += 1
            st.rerun()

# ============================
# 8. List ページ
# ============================
def list_page(sheet2):
    st.subheader("履歴一覧")

    df = sheet2.copy()
    df["date"] = pd.to_datetime(df["date"])

    weekday_map = {
        "Monday": "月", "Tuesday": "火", "Wednesday": "水",
        "Thursday": "木", "Friday": "金", "Saturday": "土", "Sunday": "日"
    }
    df["weekday"] = df["date"].dt.day_name().map(weekday_map)
    df["date_display"] = df["date"].dt.strftime("%Y-%m-%d") + "（" + df["weekday"] + "）"

    df = df.sort_values("date", ascending=False)

    display_df = df[["date_display", "type", "amount", "from", "to", "memo"]]

    st.dataframe(display_df, use_container_width=True)

# ============================
# 9. グラフページ（plotly版）
# ============================
import plotly.express as px

def graph_page(sheet1, sheet2):

    df = sheet2.copy()
    df["date"] = pd.to_datetime(df["date"])

    today = datetime.today().date()
    month_start = today.replace(day=1)

    st.header("📊 グラフ")

    # ============================
    # ① 総資産額の推移（縦棒）
    # ============================
    st.subheader("総資産額の推移（縦棒グラフ）")

    s1 = sheet1.copy()
    s1.columns = ["type", "name", "balance"]
    current_total = s1["balance"].sum()

    df2 = df[df["type"] != "振替"].copy()
    df2 = df2.sort_values("date", ascending=False)

    totals = []
    running = current_total

    for _, row in df2.iterrows():
        totals.append((row["date"].date(), running))
        running -= row["amount"]

    hist = pd.DataFrame(totals, columns=["date", "total"]).sort_values("date")

    st.bar_chart(hist.set_index("date")["total"])

    # ============================
    # ② 日別の収支（縦棒）
    # ============================
    st.subheader("日別の収支（縦棒グラフ）")

    df_income = df[df["type"] != "振替"].copy()
    daily = df_income.groupby(df_income["date"].dt.date)["amount"].sum()

    st.bar_chart(daily)

    # ============================
    # ③ 支出割合（円グラフ）
    # ④ 収入割合（円グラフ）
    # → 横並び 1段2列（plotly版）
    # ============================
    st.subheader("今月の支出・収入の割合")

    col1, col2 = st.columns(2)

    # ----------------------------
    # 支出割合（費目別 to）
    # ----------------------------
    with col1:
        st.write("今月の支出割合（費目別）")

        df_exp = df[(df["type"] == "支出") & (df["date"].dt.date >= month_start)]

        if len(df_exp) > 0:
            exp_group = df_exp.groupby("to")["amount"].sum().abs()
            fig = px.pie(
                names=exp_group.index,
                values=exp_group.values,
                title="支出割合",
                width=300,
                height=300
            )
            st.plotly_chart(fig, use_container_width=False)
        else:
            st.write("データなし")

    # ----------------------------
    # 収入割合（収入元別 from）
    # ----------------------------
    with col2:
        st.write("今月の収入割合（収入元別）")

        df_inc = df[(df["type"] == "収入") & (df["date"].dt.date >= month_start)]

        if len(df_inc) > 0:
            inc_group = df_inc.groupby("from")["amount"].sum()
            fig = px.pie(
                names=inc_group.index,
                values=inc_group.values,
                title="収入割合",
                width=300,
                height=300
            )
            st.plotly_chart(fig, use_container_width=False)
        else:
            st.write("データなし")


# ============================
# 10. ログイン処理
# ============================
auth_result = get_token(show_login_ui=False)

if not auth_result:
    st.title("資産管理")
    login_url = get_token(show_login_ui=True)

    st.markdown(
        f"""
        <a href="{login_url}">
            <button style="
                padding: 10px 20px;
                font-size: 18px;
                background-color: #4a90e2;
                color: white;
                border: none;
                border-radius: 6px;
                cursor: pointer;
            ">
                Microsoft にログイン
            </button>
        </a>
        """,
        unsafe_allow_html=True,
    )
    st.stop()

token = auth_result
st.session_state.token = token

# ============================
# 11. OneDrive 読み込み
# ============================
if "sheets" not in st.session_state:
    st.session_state.sheets = read_workbook_from_onedrive(token, FILE_PATH)

sheets = st.session_state.sheets
sheet1 = sheets["Sheet1"]
sheet2 = clean_sheet2(sheets["Sheet2"])

# ============================
# 12. メニュー
# ============================
if "page" not in st.session_state:
    st.session_state.page = "Dashboard"

menu = st.radio("メニュー", ["Dashboard", "Input", "List", "Graph"], horizontal=True)
st.session_state.page = menu

# ============================
# 13. ページ切り替え
# ============================
if st.session_state.page == "Dashboard":
    dashboard_page(sheet1, sheet2)

elif st.session_state.page == "Input":
    input_page(sheet1, sheet2, sheets, token)

elif st.session_state.page == "List":
    list_page(sheet2)

elif st.session_state.page == "Graph":
    graph_page(sheet1, sheet2)
