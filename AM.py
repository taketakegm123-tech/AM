import streamlit as st
import msal
import requests
import pandas as pd
import io

# ============================
# 1. Secrets 読み込み
# ============================
CLIENT_ID = st.secrets["azure"]["client_id"]
CLIENT_SECRET = st.secrets["azure"]["client_secret"]
TENANT_ID = st.secrets["azure"]["tenant_id"]
REDIRECT_URI = st.secrets["onedrive"]["redirect_uri"]

# OneDrive 上のファイルパス
FILE_PATH = "Asset_Manager/assets.xlsx"

# スコープ（MSA は Files.ReadWrite のみ）
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
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET,
        token_cache=cache,
    )

def get_token():
    cache = load_cache()
    app = build_msal_app(cache)

    # 既存トークンがあれば再利用
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPE, account=accounts[0])
        if result and "access_token" in result:
            return result["access_token"]

    # 初回認証
    auth_url = app.get_authorization_request_url(SCOPE, redirect_uri=REDIRECT_URI)
    st.markdown(f"[ここをクリックして Microsoft にログインする]({auth_url})")

    # Streamlit の確実に動くクエリ取得
    query_params = st.experimental_get_query_params()
    if "code" in query_params:
        code = query_params["code"][0]

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
# 3. Excel 読み込み／書き込み（複数シート対応）
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
# 4. Streamlit UI
# ============================
st.title("クラウド資産管理アプリ（OneDrive 連携版）")

token = get_token()

if not token:
    st.info("上のリンクからログインしてください。")
    st.stop()

st.success("ログイン成功！ OneDrive の Excel を読み込みます。")

sheets = read_workbook_from_onedrive(token, FILE_PATH)

if sheets is None:
    st.stop()

# シート選択
sheet_names = list(sheets.keys())
selected_sheet = st.selectbox("編集するシートを選択", sheet_names)

df = sheets[selected_sheet]
st.subheader(f"現在のデータ（{selected_sheet}）")
st.dataframe(df)

# ============================
# 仕訳追加フォーム
# ============================
st.subheader("新しい仕訳を追加")

col1, col2 = st.columns(2)
with col1:
    date = st.date_input("日付")
with col2:
    amount = st.number_input("金額", step=100)

category = st.text_input("カテゴリ")
memo = st.text_input("メモ")

if st.button("このシートに行を追加して保存"):
    new_row = {
        "日付": date,
        "金額": amount,
        "カテゴリ": category,
        "メモ": memo,
    }

    # 既存列に合わせて不足分は追加
    for col in df.columns:
        if col not in new_row:
            new_row[col] = None

    df_updated = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    sheets[selected_sheet] = df_updated

    ok = write_workbook_to_onedrive(token, FILE_PATH, sheets)

    if ok:
        st.success("OneDrive の Excel に保存しました。")
        st.experimental_rerun()
