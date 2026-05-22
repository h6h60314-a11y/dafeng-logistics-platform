from __future__ import annotations

import pandas as pd
import streamlit as st

from common_ui import card_close, card_open, inject_logistics_theme, preview_table, set_page
from google_sheet_store import ensure_database, is_configured, list_result_worksheets, read_sheet


st.set_page_config(page_title="資料庫後台", page_icon="🗄️", layout="wide")
inject_logistics_theme()

set_page(
    "資料庫後台",
    icon="🗄️",
    subtitle="以 Google Sheet 保存上傳紀錄、分析執行紀錄與結果資料。",
)


def _setup_help():
    card_open("Google Sheet 設定")
    st.warning("尚未完成 Google Sheet 連線設定。")
    st.markdown(
        """
請先建立一份 Google Sheet，並建立 Google Cloud Service Account。
將該 Service Account 的 email 加入 Google Sheet 共用權限，角色選「編輯者」。

接著在 Streamlit secrets 加入以下設定：
"""
    )
    st.code(
        """
GSHEET_ID = "你的 Google Sheet ID"

[google_service_account]
type = "service_account"
project_id = "..."
private_key_id = "..."
private_key = "-----BEGIN PRIVATE KEY-----\\n...\\n-----END PRIVATE KEY-----\\n"
client_email = "你的-service-account@專案.iam.gserviceaccount.com"
client_id = "..."
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "..."
universe_domain = "googleapis.com"
""".strip(),
        language="toml",
    )
    st.info("完成後重新整理頁面，系統會自動建立 uploads、runs 與各結果工作表。")
    card_close()


if not is_configured():
    _setup_help()
    st.stop()


card_open("連線狀態")
try:
    ensure_database()
    st.success("Google Sheet 已連線，uploads / runs 工作表已完成初始化。")
except Exception as exc:
    st.error(f"Google Sheet 初始化失敗：{type(exc).__name__}: {exc!r}")
    card_close()
    st.stop()
card_close()

tabs = st.tabs(["上傳紀錄", "分析紀錄", "結果資料表"])

with tabs[0]:
    try:
        df_uploads = read_sheet("uploads", max_rows=500)
    except Exception as exc:
        st.error(f"讀取 uploads 失敗：{type(exc).__name__}: {exc!r}")
        df_uploads = pd.DataFrame()
    preview_table("最近上傳紀錄", df_uploads, rows=500, height=460)

with tabs[1]:
    try:
        df_runs = read_sheet("runs", max_rows=500)
    except Exception as exc:
        st.error(f"讀取 runs 失敗：{type(exc).__name__}: {exc!r}")
        df_runs = pd.DataFrame()
    preview_table("最近分析紀錄", df_runs, rows=500, height=460)

with tabs[2]:
    try:
        sheets = list_result_worksheets()
    except Exception as exc:
        st.error(f"讀取結果工作表失敗：{type(exc).__name__}: {exc!r}")
        sheets = []

    if not sheets:
        st.info("目前尚未保存任何結果資料表。")
    else:
        selected = st.selectbox("選擇結果資料表", sheets)
        try:
            df_result = read_sheet(selected, max_rows=500)
        except Exception as exc:
            st.error(f"讀取 {selected} 失敗：{type(exc).__name__}: {exc!r}")
            df_result = pd.DataFrame()
        preview_table(f"{selected} 最近資料", df_result, rows=500, height=520)
