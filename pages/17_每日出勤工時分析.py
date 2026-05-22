# pages/17_每日出勤工時分析.py
# -*- coding: utf-8 -*-

import warnings
warnings.filterwarnings("ignore")

import io
import os
from datetime import date

import numpy as np
import pandas as pd
import streamlit as st

try:
    from common_ui import inject_logistics_theme, set_page, card_open, card_close
    HAS_COMMON_UI = True
except Exception:
    HAS_COMMON_UI = False


ROLE_ORDER = ["幹部", "理貨人員", "計時", "派遣", "支援本倉", "支援外倉"]
EXCLUDE_ROLE_REGEX = r"主管"
NAME_SUFFIX_STRIP_REGEX = r"\s*-(?:1|2)\s*$"
TOP_NOTE = "備註：姓名以去尾碼(-1/-2)後去重；已排除職務含『主管』；僅計工時>0"
SHEET_NAME = "總明細"


def _spacer(h: int = 10):
    st.markdown(f"<div style='height:{h}px'></div>", unsafe_allow_html=True)


def detect_role_column(cols) -> str | None:
    if "職務" in cols:
        return "職務"
    if "職務別" in cols:
        return "職務別"
    return None


def to_num(s):
    return pd.to_numeric(s, errors="coerce")


def compute_hours(df: pd.DataFrame) -> pd.Series:
    """上班時數 → 打卡時數 → (下班-上班)-用餐"""
    h = pd.Series(np.nan, index=df.index, dtype="float64")

    if "上班時數" in df.columns:
        h = to_num(df["上班時數"])

    if h.isna().all() and "打卡時數" in df.columns:
        h = to_num(df["打卡時數"])

    if h.isna().all():
        if ("上班打卡時間" in df.columns) and ("下班打卡時間" in df.columns):
            tin = pd.to_datetime(df["上班打卡時間"], errors="coerce")
            tout = pd.to_datetime(df["下班打卡時間"], errors="coerce")
            dur = (tout - tin).dt.total_seconds() / 3600.0
            meal = to_num(df.get("用餐時數", 0)).fillna(0.0)
            h = dur - meal

    return pd.to_numeric(h, errors="coerce").fillna(0.0)


def normalize_role(s: pd.Series) -> pd.Series:
    s = s.astype(str).str.strip()
    return s.replace({"": "未填", "nan": "未填", "None": "未填"})


def robust_read_excel(uploaded_file, sheet_name: str) -> pd.DataFrame:
    filename = uploaded_file.name
    ext = os.path.splitext(filename)[1].lower()
    data = uploaded_file.getvalue()
    bio = io.BytesIO(data)

    if ext in (".xlsx", ".xlsm", ".xltx", ".xltm"):
        return pd.read_excel(bio, sheet_name=sheet_name, engine="openpyxl")

    if ext == ".xls":
        return pd.read_excel(bio, sheet_name=sheet_name, engine="xlrd")

    return pd.read_excel(bio, sheet_name=sheet_name, engine="openpyxl")


def build_output_excel_bytes(
    role_counts: pd.DataFrame,
    total_headcount: int,
    hours_summary: pd.DataFrame,
    target_date: date,
    out_name: str,
) -> tuple[str, bytes]:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        sheet = "當日_各職務_工時"
        start_row = 8
        hours_summary.to_excel(writer, sheet_name=sheet, index=False, startrow=start_row)

        ws = writer.sheets[sheet]
        wb = writer.book

        big = wb.add_format({"bold": True, "font_size": 12})
        label = wb.add_format({"bold": True})
        gray = wb.add_format({"font_color": "#666666"})

        ws.write(0, 0, f"{target_date}", big)
        ws.write(1, 0, TOP_NOTE, gray)
        ws.write(2, 0, "總人次：", label)
        ws.write(2, 1, int(total_headcount), big)

        row = 3
        for _, r in role_counts.iterrows():
            ws.write(row, 0, f"{r['職務']}：", label)
            ws.write(row, 1, int(r["人次"]))
            row += 1

        ws.set_column(0, 0, 16)
        ws.set_column(1, 1, 12)
        ws.set_column(2, 10, 14)
        ws.freeze_panes(start_row + 1, 0)

    output.seek(0)
    return out_name, output.getvalue()


# =========================
# UI
# =========================
st.set_page_config(page_title="每日出勤工時分析", page_icon="🕒", layout="wide")

if HAS_COMMON_UI:
    inject_logistics_theme()
    set_page("每日出勤工時分析", icon="🕒", subtitle="出勤人次｜工時彙總｜Excel匯出")
else:
    st.title("🕒 每日出勤工時分析")

st.markdown("上傳出勤檔案（需含「總明細」分頁）並選擇日期")
_spacer(10)

# ✅ 直向：出勤Excel → 計算日期
if HAS_COMMON_UI:
    card_open("📤 出勤 Excel")
uploaded = st.file_uploader("上傳出勤 Excel（需含「總明細」分頁）", type=["xlsx", "xls", "xlsm"])
if HAS_COMMON_UI:
    card_close()

_spacer(14)

if HAS_COMMON_UI:
    card_open("📅 計算日期")
target_date = st.date_input("選擇要計算的日期", value=None)
if HAS_COMMON_UI:
    card_close()

_spacer(14)

if not uploaded:
    st.info("請先上傳出勤 Excel 檔。")
    st.stop()

if not target_date:
    st.info("請選擇要計算的日期。")
    st.stop()

# 讀檔
try:
    df = robust_read_excel(uploaded, sheet_name=SHEET_NAME)
except Exception as e:
    st.error(f"讀取失敗：找不到分頁「{SHEET_NAME}」或檔案格式不支援。\n\n錯誤訊息：{e}")
    st.stop()

if "年月日" not in df.columns:
    st.error("欄位缺少：找不到「年月日」欄位。")
    st.stop()

role_col = detect_role_column(df.columns)
if not role_col:
    st.error("欄位缺少：找不到「職務」或「職務別」欄位。")
    st.stop()

if "員工姓名" not in df.columns:
    st.error("欄位缺少：找不到「員工姓名」欄位。")
    st.stop()

# 計算
df["日期"] = pd.to_datetime(df["年月日"], errors="coerce").dt.date
df["工時"] = compute_hours(df)

day = df[df["日期"] == target_date].copy()
if day.empty:
    st.warning(f"{target_date} 無出勤資料。")
    st.stop()

day[role_col] = normalize_role(day[role_col])
day["員工姓名"] = day["員工姓名"].astype(str).str.strip()

day = day[~day[role_col].str.contains(EXCLUDE_ROLE_REGEX, na=False)].copy()
day = day[day["工時"] > 0].copy()

day["姓名_去尾碼"] = day["員工姓名"].str.replace(NAME_SUFFIX_STRIP_REGEX, "", regex=True).str.strip()
total_headcount = int(day["姓名_去尾碼"].nunique())

role_counts = (
    day.groupby(role_col)["姓名_去尾碼"]
       .nunique()
       .reindex(ROLE_ORDER, fill_value=0)
       .reset_index()
       .rename(columns={role_col: "職務", "姓名_去尾碼": "人次"})
)

hours_summary = (
    day.groupby(role_col)["工時"].sum()
       .reindex(ROLE_ORDER, fill_value=0)
       .reset_index()
)
hours_summary.columns = ["職務", "工時"]
hours_summary = pd.concat(
    [hours_summary, pd.DataFrame([{"職務": "總計", "工時": hours_summary["工時"].sum()}])],
    ignore_index=True
)
hours_summary["工時"] = hours_summary["工時"].round(2)

counts_map = {r["職務"]: int(r["人次"]) for _, r in role_counts.iterrows()}
hours_map = {r["職務"]: float(r["工時"]) for _, r in hours_summary.iterrows()}

_spacer(10)

# ✅✅ 把備註移到兩欄上方（卡片外），避免左欄往下推造成不對齊
st.caption(TOP_NOTE)
_spacer(6)

# =========================
# 左右兩欄：直向清單（✅總人次與右欄幹部同高）
# =========================
left, right = st.columns([1, 1], gap="large")

with left:
    if HAS_COMMON_UI:
        card_open("👥 當日人次總覽")
    else:
        st.subheader("👥 當日人次總覽")

    st.metric("總人次（去尾碼去重）", f"{total_headcount:,}")
    _spacer(6)
    for role in ROLE_ORDER:
        st.metric(role, f"{counts_map.get(role, 0):,}")

    if HAS_COMMON_UI:
        card_close()

with right:
    if HAS_COMMON_UI:
        card_open("🧾 各職務總上班時間（小時）")
    else:
        st.subheader("🧾 各職務總上班時間（小時）")

    for role in ROLE_ORDER:
        st.metric(role, f"{hours_map.get(role, 0.0):,.2f}")
    _spacer(6)
    st.metric("總計", f"{hours_map.get('總計', 0.0):,.2f}")

    if HAS_COMMON_UI:
        card_close()

_spacer(14)

# 下載
base = os.path.splitext(uploaded.name)[0]
out_name = f"{base}_{target_date}_工時與人次.xlsx"

download_name, excel_bytes = build_output_excel_bytes(
    role_counts=role_counts,
    total_headcount=total_headcount,
    hours_summary=hours_summary,
    target_date=target_date,
    out_name=out_name,
)

st.download_button(
    label="⬇️ 下載 Excel（工時與人次）",
    data=excel_bytes,
    file_name=download_name,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)
