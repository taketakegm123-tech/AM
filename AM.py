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
        auth_url = app.get_authorization_request_url(SCOPE, redirect_uri=REDIRECT_URI)
        return auth_url

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
# 5. Dashboard UI
# ============================
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
        .subtitle {
            font-size: 22px !important;
            font-weight: 600 !important;
            color: #555 !important;
            margin: 12px 0 6px 0 !important;
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
            color: #555 !important;
        }
        .big-card h2 {
            font-size: 14px !important;
            margin: 0 !important;
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

    st.markdown("<div class='subtitle'>カテゴリ別資産</div>", unsafe_allow_html=True)

    col_bank, col_cash, col_invest = st.columns(3)

    for col, t in zip([col_bank, col_cash, col_invest], ["bank", "cash", "invest"]):
        with col:
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

    st.markdown("<div class='subtitle'>資産推移（概算）</div>", unsafe_allow_html=True)
    st.line_chart(history.set_index("date")["total"])

# ============================
# 6. メイン UI（固定ヘッダー＋ログイン前後）
# ============================
auth_result = get_token(show_login_ui=False)

# ★★★ ログイン前画面 ★★★
if not auth_result:
    st.markdown(
        """
        <style>
        .login-title {
            font-size: 26px !important;
            font-weight: 700 !important;
            color: #555 !important;
            margin-bottom: 20px !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("<div class='login-title'>資産管理</div>", unsafe_allow_html=True)

    login_url = get_token(show_login_ui=True)

    if st.button("Microsoft にログイン"):
        st.markdown(f"[ログインはこちら]({login_url})", unsafe_allow_html=True)

    st.stop()

# ★★★ ログイン後 ★★★
token = auth_result

# 固定ヘッダー
st.markdown(
    """
    <style>
    .fixed-header {
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        background-color: #f5f5f5;
        padding: 10px 0 5px 0;
        z-index: 9999;
        text-align: center;
        border-bottom: 1px solid #ddd;
    }
    .header-title {
        font-size: 26px !important;
        font-weight: 700 !important;
        color: #555 !important;
        margin-bottom: 6px !important;
    }
    .content {
        margin-top: 120px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="fixed-header">
        <div class="header-title">資産管理メニュー</div>
    </div>
    """,
    unsafe_allow_html=True,
)

# メニュー（固定ヘッダーの下に表示）
menu = st.radio(
    "",
    ["🏠 Dashboard", "➕ Input", "📄 List", "📊 Charts"],
    horizontal=True,
)

st.markdown('<div class="content">', unsafe_allow_html=True)

# OneDrive 読み込み
sheets = read_workbook_from_onedrive(token, FILE_PATH)
if sheets is None:
    st.stop()

sheet1 = sheets["Sheet1"]
sheet2 = sheets["Sheet2"]

# メニュー分岐
if menu == "🏠 Dashboard":
    dashboard(sheet1, sheet2)

elif menu == "➕ Input":
    st.markdown("<div class='subtitle'>新しい仕訳を追加</div>", unsafe_allow_html=True)

    sheet2 = sheet2.copy()
    sheet2.columns = ["date", "from", "to", "amount", "memo"]

    col1, col2 = st.columns(2)
    with col1:
        date = st.date_input("日付")
    with col2:
        amount = st.number_input("金額", step=100)

    from_ = st.text_input("from（出金元）")
    to_ = st.text_input("to（入金先／費目）")
    memo = st.text_input("メモ")

    if st.button("仕訳を追加して保存"):
        new_row = {
            "date": date,
            "from": from_,
            "to": to_,
            "amount": amount,
            "memo": memo,
        }
        df_updated = pd.concat([sheet2, pd.DataFrame([new_row])], ignore_index=True)
        sheets["Sheet2"] = df_updated
        ok = write_workbook_to_onedrive(token, FILE_PATH, sheets)
        if ok:
            st.success("OneDrive の Excel に保存しました。")
            st.experimental_rerun()

elif menu == "📄 List":
    st.markdown("<div class='subtitle'>履歴一覧</div>", unsafe_allow_html=True)
    sheet2 = sheet2.copy()
    sheet2.columns = ["date", "from", "to", "amount", "memo"]
    st.dataframe(sheet2)

elif menu == "📊 Charts":
    st.markdown("<div class='subtitle'>金額の推移</div>", unsafe_allow_html=True)
    sheet2 = sheet2.copy()
    sheet2.columns = ["date", "from", "to", "amount", "memo"]
    sheet2["date"] = pd.to_datetime(sheet2["date"])
    st.line_chart(sheet2.set_index("date")["amount"])

st.markdown("</div>", unsafe_allow_html=True)
