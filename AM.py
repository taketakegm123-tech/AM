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

def get_token():
    cache = load_cache()
    app = build_msal_app(cache)

    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPE, account=accounts[0])
        if result and "access_token" in result:
            return result["access_token"]

    auth_url = app.get_authorization_request_url(SCOPE, redirect_uri=REDIRECT_URI)
    st.markdown(f"[ここをクリックして Microsoft にログインする]({auth_url})")

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

    if response.status_code in [200, 201]:
        return True
    else:
        st.error("書き込みに失敗しました")
        st.write(response.text)
        return False

# ============================
# 4. Dashboard 計算ロジック
# ============================
def calc_total_and_history(sheet1, sheet2):
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
        last_date = daily_delta.index.max().date()
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
# 5. スマホ最適化 Dashboard UI（完全版）
# ============================
def dashboard(sheet1, sheet2):

    TYPE_JP = {
        "bank": "銀行",
        "cash": "現金",
        "invest": "投資"
    }

    st.markdown(
        """
        <style>
        .big-card {
            padding: 12px;
            border-radius: 12px;
            background-color: #0f172a;
            color: white;
            box-shadow: 0 4px 12px rgba(0,0,0,0.3);
            text-align: center;
            margin-bottom: 16px;
        }
        .big-card h1 { font-size: 26px; margin: 0; }
        .big-card h2 { font-size: 16px; margin: 2px 0; }

        .cat-title {
            font-weight: bold;
            font-size: 18px;
            margin-bottom: 4px;
        }
        .cat-box {
            background: #1e293b;
            padding: 10px;
            border-radius: 8px;
            color: white;
            margin-bottom: 8px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    current_total, history = calc_total_and_history(sheet1, sheet2)
    diff_day, diff_month = calc_diff(current_total, history)

    st.title("資産管理")

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

    st.subheader("カテゴリ別資産")

    col_bank, col_cash, col_invest = st.columns(3)

    for col, t in zip([col_bank, col_cash, col_invest], ["bank", "cash", "invest"]):
        with col:
            st.markdown(f"<div class='cat-title'>{TYPE_JP[t]}</div>", unsafe_allow_html=True)

            df_cat = sheet1[sheet1["type"] == t]

            for _, row in df_cat.iterrows():
                st.markdown(
                    f"""
                    <div class="cat-box">
                        {row['name']}<br>
                        ¥{row['balance']:,.0f}
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

    st.subheader("資産推移（概算）")
    st.line_chart(history.set_index("date")["total"])

# ============================
# 6. メイン UI
# ============================
st.title("資産管理")

token = get_token()
if not token:
    st.info("上のリンクからログインしてください。")
    st.stop()

sheets = read_workbook_from_onedrive(token, FILE_PATH)
if sheets is None:
    st.stop()

sheet1 = sheets["Sheet1"]
sheet2 = sheets["Sheet2"]

menu = st.radio(
    "メニュー",
    ["🏠 Dashboard", "➕ Input", "📄 List", "📊 Charts"],
    horizontal=True,
)

if menu == "🏠 Dashboard":
    dashboard(sheet1, sheet2)

elif menu == "➕ Input":
    st.write("（後でスマホ最適化版にする）")
    st.dataframe(sheet2)

elif menu == "📄 List":
    st.dataframe(sheet2)

elif menu == "📊 Charts":
    st.line_chart(sheet2.set_index("date")["amount"])
