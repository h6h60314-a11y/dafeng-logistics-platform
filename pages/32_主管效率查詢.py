from __future__ import annotations

import datetime as dt

import pandas as pd
import streamlit as st

from common_ui import card_close, card_open, inject_logistics_theme, preview_table, render_kpis, set_page, KPI
from google_sheet_store import ensure_database, is_configured, read_efficiency_daily, service_account_email


st.set_page_config(page_title="主管效率查詢", page_icon="📊", layout="wide")
inject_logistics_theme()

set_page(
    "主管效率查詢",
    icon="📊",
    subtitle="依日期區間查詢全體作業效率、達標率與人員明細。",
)


def _num(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0)


def _to_bool(series: pd.Series) -> pd.Series:
    return series.astype(str).str.lower().isin(["true", "1", "yes", "y", "達標"])


if not is_configured():
    card_open("Google Sheet 尚未設定")
    st.warning("請先到 Streamlit Secrets 設定 GSHEET_ID 與 google_service_account。")
    card_close()
    st.stop()

try:
    ensure_database()
except Exception as exc:
    card_open("Google Sheet 權限不足")
    st.error(f"無法初始化主管效率資料表：{type(exc).__name__}: {exc!r}")
    email = service_account_email()
    if email:
        st.info(f"請把這個 service account 加到 Google Sheet 共用，權限選編輯者：{email}")
    card_close()
    st.stop()


today = dt.date.today()
card_open("查詢條件")
c1, c2, c3 = st.columns([1, 1, 1.2])
with c1:
    start_date = st.date_input("開始日期", value=today)
with c2:
    end_date = st.date_input("結束日期", value=today)
with c3:
    keyword = st.text_input("人員 / 作業關鍵字", value="", placeholder="可輸入姓名、作業別、部門")
card_close()

if start_date > end_date:
    st.error("開始日期不可晚於結束日期。")
    st.stop()

try:
    df = read_efficiency_daily(start_date, end_date, max_rows=10000)
except Exception as exc:
    st.error(f"讀取主管效率資料失敗：{type(exc).__name__}: {exc!r}")
    st.stop()

if df is None or df.empty:
    st.info("這個日期區間目前沒有主管效率資料。請先在效率分析頁執行分析，系統會寫入 efficiency_daily。")
    st.stop()

for col in ["department", "operation", "employee_id", "employee_name", "shift", "source_page"]:
    if col not in df.columns:
        df[col] = ""

if keyword.strip():
    key = keyword.strip()
    mask = (
        df["department"].astype(str).str.contains(key, case=False, na=False)
        | df["operation"].astype(str).str.contains(key, case=False, na=False)
        | df["employee_id"].astype(str).str.contains(key, case=False, na=False)
        | df["employee_name"].astype(str).str.contains(key, case=False, na=False)
    )
    df = df[mask].copy()

if df.empty:
    st.info("符合條件的資料為空。")
    st.stop()

df["work_count"] = _num(df.get("work_count", pd.Series(dtype=float)))
df["work_minutes"] = _num(df.get("work_minutes", pd.Series(dtype=float)))
df["efficiency"] = _num(df.get("efficiency", pd.Series(dtype=float)))
df["_has_pass"] = df.get("is_pass", pd.Series(dtype=str)).astype(str).str.strip().ne("")
df["_pass"] = _to_bool(df.get("is_pass", pd.Series(dtype=str)))

total_count = float(df["work_count"].sum())
total_minutes = float(df["work_minutes"].sum())
weighted_eff = (total_count / total_minutes * 60) if total_minutes > 0 else float(df["efficiency"].mean())
people = int(df[["employee_id", "employee_name"]].drop_duplicates().shape[0])
pass_rate = float(df.loc[df["_has_pass"], "_pass"].mean()) if bool(df["_has_pass"].any()) else 0.0

card_open("整體效率")
render_kpis(
    [
        KPI("查詢筆數", f"{len(df):,}"),
        KPI("人數", f"{people:,}"),
        KPI("總作業量", f"{total_count:,.0f}"),
        KPI("總工時(小時)", f"{total_minutes / 60:,.2f}"),
        KPI("整體效率", f"{weighted_eff:,.2f}"),
        KPI("達標率", f"{pass_rate:.0%}"),
    ]
)
card_close()

person_group_cols = ["employee_id", "employee_name"]
person = (
    df.groupby(person_group_cols, dropna=False)
    .agg(
        作業量=("work_count", "sum"),
        工時分鐘=("work_minutes", "sum"),
        平均效率=("efficiency", "mean"),
        紀錄筆數=("efficiency", "size"),
        達標率=("_pass", "mean"),
        已判定達標筆數=("_has_pass", "sum"),
    )
    .reset_index()
)
person["整體效率"] = person.apply(
    lambda r: (float(r["作業量"]) / float(r["工時分鐘"]) * 60) if float(r["工時分鐘"]) > 0 else float(r["平均效率"]),
    axis=1,
)
person["工時小時"] = person["工時分鐘"] / 60
person = person.rename(columns={"employee_id": "人員代碼", "employee_name": "姓名"})
person.loc[person["已判定達標筆數"] == 0, "達標率"] = None
person = person[["人員代碼", "姓名", "作業量", "工時小時", "整體效率", "平均效率", "達標率", "紀錄筆數"]]
person = person.sort_values("整體效率", ascending=False)

op = (
    df.groupby(["department", "operation"], dropna=False)
    .agg(
        作業量=("work_count", "sum"),
        工時分鐘=("work_minutes", "sum"),
        平均效率=("efficiency", "mean"),
        人數=("employee_name", "nunique"),
        達標率=("_pass", "mean"),
        已判定達標筆數=("_has_pass", "sum"),
    )
    .reset_index()
)
op["整體效率"] = op.apply(
    lambda r: (float(r["作業量"]) / float(r["工時分鐘"]) * 60) if float(r["工時分鐘"]) > 0 else float(r["平均效率"]),
    axis=1,
)
op["工時小時"] = op["工時分鐘"] / 60
op = op.rename(columns={"department": "部門", "operation": "作業"})
op.loc[op["已判定達標筆數"] == 0, "達標率"] = None
op = op[["部門", "作業", "人數", "作業量", "工時小時", "整體效率", "平均效率", "達標率"]]
op = op.sort_values(["部門", "作業"])

preview_table("人員效率排行", person, rows=500, height=420)
preview_table("作業別整體效率", op, rows=500, height=320)
preview_table("原始明細", df.drop(columns=["_pass"], errors="ignore"), rows=1000, height=460)
