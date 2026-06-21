import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import bcrypt
import uuid
from datetime import datetime
import json
import os

# =========================================================
# Google Sheets 接続（JSONファイルを直接読み込む）
# =========================================================
def get_sheet():
    # JSON キーのファイル名（GitHub にアップした名前）
    json_path = "asset-manager-500108-fbfdacc57942.json"

    # JSON を読み込む
    with open(json_path, "r") as f:
        creds_info = json.load(f)

    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_info(creds_info, scopes=scopes)

    client = gspread.authorize(creds)

    # スプレッドシートIDは Secrets から読み込む（これは短いので壊れにくい）
    sheet_id = st.secrets["sheet_id"]

    sheet = client.open_by_key(sheet_id).sheet1
    return sheet

# =========================================================
# パスワード管理（session_state に保持）
# =========================================================
def load_passwords():
    if "passwords" not in st.session_state:
        st.session_state["passwords"] = {
            "YT": bcrypt.hashpw("PW123".encode(), bcrypt.gensalt()).decode(),
            "Guest": bcrypt.hashpw("PW12345".encode(), bcrypt.gensalt()).decode()
        }
    return st.session_state["passwords"]

def save_passwords(passwords):
    st.session_state["passwords"] = passwords

# =========================================================
# 家計簿データ読み込み
# =========================================================
def load_kakeibo():
    sheet = get_sheet()
    rows = sheet.get_all_records()
    return rows

# =========================================================
# 家計簿データ保存（1行追加）
# =========================================================
def save_kakeibo_row(row):
    sheet = get_sheet()
    sheet.append_row(row)

# =========================================================
# ログイン画面
# =========================================================
def login_screen():
    st.title("🔐 ログイン")

    username = st.text_input("ユーザー名")
    password = st.text_input("パスワード", type="password")

    if st.button("ログイン"):
        passwords = load_passwords()

        if username in passwords:
            hashed = passwords[username].encode()
            if bcrypt.checkpw(password.encode(), hashed):
                st.session_state["user"] = username
                st.session_state["page"] = "main"
                st.rerun()
            else:
                st.error("パスワードが違います")
        else:
            st.error("ユーザー名が存在しません")

# =========================================================
# パスワード変更画面
# =========================================================
def change_password_screen():
    st.title("🔑 パスワード変更")

    old_pw = st.text_input("現在のパスワード", type="password")
    new_pw = st.text_input("新しいパスワード", type="password")

    if st.button("変更する"):
        passwords = load_passwords()
        username = st.session_state["user"]

        if bcrypt.checkpw(old_pw.encode(), passwords[username].encode()):
            passwords[username] = bcrypt.hashpw(new_pw.encode(), bcrypt.gensalt()).decode()
            save_passwords(passwords)
            st.success("パスワードを変更しました")
            st.session_state["page"] = "main"
            st.rerun()
        else:
            st.error("現在のパスワードが違います")

# =========================================================
# 家計簿入力画面
# =========================================================
def kakeibo_input_screen():
    st.title("📘 家計簿入力")

    date = st.date_input("日付", datetime.now())
    category = st.selectbox("カテゴリ", ["食費", "交通費", "日用品", "収入", "その他"])
    amount = st.number_input("金額", step=100, format="%d")
    memo = st.text_input("メモ（任意）")

    if st.button("追加"):
        row = [
            str(uuid.uuid4()),
            str(date),
            category,
            int(amount),
            memo
        ]
        save_kakeibo_row(row)
        st.success("追加しました！")

# =========================================================
# 家計簿一覧画面
# =========================================================
def kakeibo_list_screen():
    st.title("📄 家計簿一覧")

    data = load_kakeibo()
    if not data:
        st.info("まだデータがありません")
        return

    st.table(data)

# =========================================================
# ページ描画
# =========================================================
def render_page():
    page = st.session_state.get("page", "login")

    if page == "login":
        login_screen()
    elif page == "main":
        kakeibo_input_screen()
    elif page == "list":
        kakeibo_list_screen()
    elif page == "change_pw":
        change_password_screen()

# =========================================================
# メニュー（常時表示）
# =========================================================
if "user" in st.session_state:
    menu = st.sidebar.radio(
        "メニュー",
        ["家計簿入力", "家計簿一覧", "パスワード変更", "ログアウト"]
    )

    if menu == "家計簿入力":
        st.session_state["page"] = "main"
        st.rerun()
    elif menu == "家計簿一覧":
        st.session_state["page"] = "list"
        st.rerun()
    elif menu == "パスワード変更":
        st.session_state["page"] = "change_pw"
        st.rerun()
    elif menu == "ログアウト":
        st.session_state.clear()
        st.session_state["page"] = "login"
        st.rerun()

# =========================================================
# ページ描画
# =========================================================
render_page()
