import streamlit as st
import msal
import requests
import pandas as pd
import json

# ============================
# 1. Streamlit Secrets 読み込み
# ============================
CLIENT_ID = st.secrets["azure"]["client_id"]
CLIENT_SECRET = st.secrets["azure"]["client_secret"]
TENANT_ID = st.secrets["azure"]["tenant_id"]
REDIRECT_URI = st.secrets["onedrive"]["redirect_uri"]

# Microsoft Graph のスコープ
SCOPE = ["Files.ReadWrite", "User.Read", "offline_access"]

# ============================
# 2. 認証用の MSAL オブジェクト
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
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET,
        token_cache=cache
    )

# ============================
# 3. 認証フロー
# ============================
def get_token():
    cache = load_cache()
    app = build_msal_app(cache)

    # 既存トークンがあれば使う
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPE, account=accounts[0])
        if result:
            return result["access_token"]

    # 初回ログイン
    auth_url = app.get_authorization_request_url(SCOPE, redirect_uri=REDIRECT_URI)
    st.markdown(f"[ここをクリックして Microsoft にログインする]({auth_url})")

    # 認証後のコードを受け取る
    query_params = st.query_params
    if "code" in query_params:
        code = query_params["code"]
        result = app.acquire_token_by_authorization_code(
            code,
            scopes=SCOPE,
            redirect_uri=REDIRECT_URI
        )
        if "access_token" in result:
            save_cache(cache)
            return result["access_token"]

    return None

# ============================
# 4. OneDrive の Excel を読み込む関数
# ============================
def read_excel_from_onedrive(access_token, file_path):
    url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{file_path}:/content"
    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        return pd.read_excel(response.content)
    else:
        st.error("ファイル取得に失敗しました")
        st.write(response.text)
        return None

# ============================
# 5. Streamlit UI
# ============================
st.title("OneDrive Excel 読み込みテスト")

token = get_token()

if token:
    st.success("ログイン成功！Excel を読み込みます。")

    df = read_excel_from_onedrive(token, "Asset_Manager/assets.xlsx")  # ← OneDrive のパス
    if df is not None:
        st.dataframe(df)
else:
    st.info("ログインしてください。")
