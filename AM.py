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
# 4. Dashboard 計算ロジック
# ============================
def calc_total_and_history(sheet1, sheet2):
    sheet1 = sheet1.copy()
    sheet2 = sheet2.copy()
    sheet1.columns = ["type", "name", "balance"]
    sheet2.columns = ["date", "from", "to", "amount", "memo"]
    sheet2["date"] = pd.to_datetime(sheet2["date"])

    current_total = sheet1["balance"].sum()
    asset_types = sheet1["type"].unique().tolist()

    def row_delta(row):
        f = row["from"]
        t = row["to"]
        amt = row["amount"]
        f_is_asset = f in asset_types
        t_is_asset = t in asset_types

        if f_is_asset and t_is_asset:
            return 0
        if f_is_asset and not t_is_asset:
            return -amt
        if not f_is_asset and t_is_asset:
            return +amt
        return 0

    sheet2["delta"] = sheet2.apply(row_delta, axis=1)
    daily_delta = sheet2.groupby("date")["delta"].sum().sort_index()

    if len(daily_delta) == 0:
        history = pd.DataFrame({
            "date": [datetime.today().date()],
            "total": [current_total]
        })
    else:
        daily_delta_sorted = daily_delta.sort_index(ascending=False)
        totals = []
        running_total = current_total
        for d, delta in daily_delta_sorted.items():
            totals.append((d.date(), running_total))
            running_total -= delta
        history = pd.DataFrame(totals, columns=["date", "total"]).sort_values("date")

    return current_total, history

def calc_diff(current_total, history):
    today = history["date"].max()
    yesterday = today - timedelta(days=1)
    last_month = today - timedelta(days=30)

    y_row = history[history["date"] <= yesterday]
    yesterday_total = y_row["total"].iloc[-1] if len(y_row) else current_total

    m_row = history[history["date"] <= last_month]
    last_month_total = m_row["total"].iloc[-1] if len(m_row) else current_total

    return current_total - yesterday_total, current_total - last_month_total

# ============================
# 5. Dashboard ページ
# ============================
def dashboard_page(sheet1, sheet2):
    current_total, history = calc_total_and_history(sheet1, sheet2)
    diff_day, diff_month = calc_diff(current_total, history)

    st.markdown(
        """
        <style>
        .big-card {
            padding: 10px;
            border-radius: 10px;
            background-color: #e9d5ff !important;
            border: 1px solid #aaa;
            color: #555 !important;
            text-align: center;
            margin-bottom: 12px;
        }
        .big-card h1 {
            font-size: 18px !important;
            margin: 0 !important;
        }
        .big-card h2 {
            font-size: 12px !important;
            margin: 0 !important;
        }
        .subtitle {
            font-size: 22px !important;
            font-weight: 700 !important;
            color: #555 !important;
            margin: 20px 0 10px 0 !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        f"""
        <div class="big-card">
            <h1>総資産：¥{current_total:,.0f}</h1>
            <h2>前日比：{('+' if diff_day >= 0 else '')}¥{diff_day:,.0f} ｜ 前月比：{('+' if diff_month >= 0 else '')}¥{diff_month:,.0f}</h2>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("<div class='subtitle'>カテゴリ別資産</div>", unsafe_allow_html=True)

    TYPE_COLOR = {
        "bank": "#dbeafe",
        "cash": "#dcfce7",
        "invest": "#fde2e4"
    }

    col1, col2, col3 = st.columns(3)
    for col, t in zip([col1, col2, col3], ["bank", "cash", "invest"]):
        df_cat = sheet1[sheet1["type"] == t]
        for _, row in df_cat.iterrows():
            col.markdown(
                f"""
                <div style="padding:10px; border-radius:8px; border:1px solid #aaa; background-color:{TYPE_COLOR[t]}; margin-bottom:10px;">
                    {row['name']}<br>
                    ¥{row['balance']:,.0f}
                </div>
                """,
                unsafe_allow_html=True,
            )

    st.markdown("<div class='subtitle'>資産推移（概算）</div>", unsafe_allow_html=True)
    st.line_chart(history.set_index("date")["total"])

# ============================
# 6. Input ページ（支出・収入フレーム）
# ============================
def input_page(sheet1, sheet2, sheets, token):

    # 初期化
    defaults = {
        "exp_date": datetime.today().date(),
        "exp_amount": 0,
        "exp_from": "財布",
        "exp_to": "",
        "inc_date": datetime.today().date(),
        "inc_amount": 0,
        "inc_from": "",
        "inc_to": "財布",
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val

    asset_list = sheet1["name"].tolist()
    default_wallet_index = asset_list.index("財布") if "財布" in asset_list else 0

    # ============================
    # 🟥 支出フレーム
    # ============================
    with st.container(border=True):
        st.subheader("支出")

        st.session_state.exp_date = st.date_input(
            "日付",
            value=st.session_state.exp_date,
            key="exp_date_input"
        )

        st.session_state.exp_amount = st.number_input(
            "金額（プラスで入力）",
            min_value=0,
            step=100,
            value=st.session_state.exp_amount,
            key="exp_amount_input"
        )

        st.session_state.exp_from = st.selectbox(
            "from（出金元）",
            asset_list,
            index=default_wallet_index,
            key="exp_from_select"
        )

        st.session_state.exp_to = st.text_input(
            "to（費目）",
            value=st.session_state.exp_to,
            key="exp_to_input"
        )

        if st.button("支出を入力"):
            date = pd.to_datetime(st.session_state.exp_date)
            amount = -abs(st.session_state.exp_amount)

            new_row = {
                "date": date,
                "from": st.session_state.exp_from,
                "to": st.session_state.exp_to,
                "amount": amount,
                "memo": ""
            }

            sheet2_local = sheet2.copy()
            sheet2_local.columns = ["date", "from", "to", "amount", "memo"]
            df_updated = pd.concat([sheet2_local, pd.DataFrame([new_row])], ignore_index=True)

            sheets["Sheet2"] = df_updated

            if write_workbook_to_onedrive(token, FILE_PATH, sheets):
                st.session_state.sheets = read_workbook_from_onedrive(token, FILE_PATH)

                st.session_state.exp_date = datetime.today().date()
                st.session_state.exp_amount = 0
                st.session_state.exp_from = "財布"
                st.session_state.exp_to = ""

                st.rerun()

    # ============================
    # 🟦 収入フレーム
    # ============================
    with st.container(border=True):
        st.subheader("収入")

        st.session_state.inc_date = st.date_input(
            "日付",
            value=st.session_state.inc_date,
            key="inc_date_input"
        )

        st.session_state.inc_amount = st.number_input(
            "金額（プラスで入力）",
            min_value=0,
            step=100,
            value=st.session_state.inc_amount,
            key="inc_amount_input"
        )

        st.session_state.inc_from = st.text_input(
            "from（収入元）",
            value=st.session_state.inc_from,
            key="inc_from_input"
        )

        st.session_state.inc_to = st.selectbox(
            "to（入金先資産）",
            asset_list,
            index=default_wallet_index,
            key="inc_to_select"
        )

        if st.button("収入を入力"):
            date = pd.to_datetime(st.session_state.inc_date)
            amount = abs(st.session_state.inc_amount)

            new_row = {
                "date": date,
                "from": st.session_state.inc_from,
                "to": st.session_state.inc_to,
                "amount": amount,
                "memo": ""
            }

            sheet2_local = sheet2.copy()
            sheet2_local.columns = ["date", "from", "to", "amount", "memo"]
            df_updated = pd.concat([sheet2_local, pd.DataFrame([new_row])], ignore_index=True)

            sheets["Sheet2"] = df_updated

            if write_workbook_to_onedrive(token, FILE_PATH, sheets):
                st.session_state.sheets = read_workbook_from_onedrive(token, FILE_PATH)

                st.session_state.inc_date = datetime.today().date()
                st.session_state.inc_amount = 0
                st.session_state.inc_from = ""
                st.session_state.inc_to = "財布"

                st.rerun()

# ============================
# 7. List ページ
# ============================
def list_page(sheet2):
    st.subheader("履歴一覧")
    sheet2_local = sheet2.copy()
    sheet2_local.columns = ["date", "from", "to", "amount", "memo"]
    st.dataframe(sheet2_local)

# ============================
# 8. Charts ページ
# ============================
def charts_page(sheet2):
    st.subheader("金額の推移")
    sheet2_local = sheet2.copy()
    sheet2_local.columns = ["date", "from", "to", "amount", "memo"]
    sheet2_local["date"] = pd.to_datetime(sheet2_local["date"])
    st.line_chart(sheet2_local.set_index("date")["amount"])

# ============================
# 9. ログイン処理
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

# ============================
# 10. OneDrive 読み込み（高速化）
# ============================
if "sheets" not in st.session_state:
    st.session_state.sheets = read_workbook_from_onedrive(token, FILE_PATH)

sheets = st.session_state.sheets
sheet1 = sheets["Sheet1"]
sheet2 = sheets["Sheet2"]

# ============================
# 11. メニュー
# ============================
if "page" not in st.session_state:
    st.session_state.page = "Dashboard"

menu = st.radio(
    "メニュー",
    ["Dashboard", "Input", "List", "Charts"],
    horizontal=True,
)

st.session_state.page = menu

# ============================
# 12. ページ切り替え
# ============================
if st.session_state.page == "Dashboard":
    dashboard_page(sheet1, sheet2)

elif st.session_state.page == "Input":
    input_page(sheet1, sheet2, sheets, token)

elif st.session_state.page == "List":
    list_page(sheet2)

elif st.session_state.page == "Charts":
    charts_page(sheet2)
