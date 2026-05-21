# pages/25_整體作業工時.py
# -*- coding: utf-8 -*-
from __future__ import annotations

from io import BytesIO
from datetime import datetime

import pandas as pd
import streamlit as st

from common_ui import (
    inject_logistics_theme,
    set_page,
    KPI,
    render_kpis,
    card_open,
    card_close,
    download_excel_card,  # ✅ 一行=按鈕（且外框不分段）
)

st.set_page_config(page_title="大豐KPI｜整體作業工時", page_icon="🕒", layout="wide")
inject_logistics_theme()

set_page(
    "整體作業工時",
    icon="🕒",
    subtitle="出勤報表｜排除空打卡＋外倉職務｜工時摘要＋明細下載",
)

# ----------------------------
# helpers
# ----------------------------
REQ_COLS = ["上班打卡時間", "職務", "組別", "上班時數", "打卡時數", "員工姓名"]


def _fmt2(x) -> str:
    try:
        return f"{float(x):,.2f}"
    except Exception:
        return ""


def _safe_str(s: pd.Series) -> pd.Series:
    return s.astype(str).fillna("").astype(str)


def get_unique_names(series: pd.Series) -> pd.Series:
    s = _safe_str(series).str.strip()
    s = s[s != ""]
    return s.str[:3].drop_duplicates()


def robust_read_excel(uploaded_file) -> pd.DataFrame:
    raw = uploaded_file.getvalue()
    bio = BytesIO(raw)

    try:
        return pd.read_excel(bio, engine="openpyxl")
    except Exception:
        try:
            bio.seek(0)
            return pd.read_excel(bio, engine="xlrd")
        except Exception as e:
            raise RuntimeError(f"讀取 Excel 失敗：{e}")


def build_outputs(df_raw: pd.DataFrame) -> dict:
    miss = [c for c in REQ_COLS if c not in df_raw.columns]
    if miss:
        raise ValueError(f"缺少必要欄位：{', '.join(miss)}")

    df0 = df_raw.copy()

    # 3) 排除「上班打卡時間」空值/空白
    before_c = len(df0)
    mask_empty_clockin = (
        df0["上班打卡時間"].isna() | (_safe_str(df0["上班打卡時間"]).str.strip() == "")
    )
    df1 = df0[~mask_empty_clockin].copy()
    removed_c = before_c - len(df1)

    # 4) 排除「職務」外倉關鍵字
    exclude_job_keywords = ["支援外倉", "倉服部", "不在本倉加班去外倉"]
    job_pattern = "|".join(exclude_job_keywords)

    before_g = len(df1)
    mask_job = _safe_str(df1["職務"]).str.contains(job_pattern, na=False)
    df_base = df1[~mask_job].copy()
    removed_g = before_g - len(df_base)

    # 6A) 一般人員（不含行政＆幹部）
    exclude_group_general = [
        "出貨主管", "行政組", "行政/盤點",
        "行政幹部", "行政主管", "驗收幹部",
        "總揀驗收主管", "總揀幹部",
    ]
    general_pattern = "|".join(exclude_group_general)
    df_general = df_base[~_safe_str(df_base["組別"]).str.contains(general_pattern, na=False)].copy()

    # 6B) 不含幹部（行政有算）
    exclude_group_noncadre = [
        "出貨主管", "行政幹部", "行政主管",
        "驗收幹部", "總揀驗收主管", "總揀幹部",
    ]
    noncadre_pattern = "|".join(exclude_group_noncadre)
    df_noncadre = df_base[~_safe_str(df_base["組別"]).str.contains(noncadre_pattern, na=False)].copy()

    # 6C) 全體
    df_all = df_base.copy()

    # 6D) 成箱組
    df_box = df_base[_safe_str(df_base["組別"]).str.contains("成箱組", na=False)].copy()

    def _calc(df_: pd.DataFrame) -> dict:
        wh = pd.to_numeric(df_["上班時數"], errors="coerce").sum()
        ch = pd.to_numeric(df_["打卡時數"], errors="coerce").sum()
        names = get_unique_names(df_["員工姓名"])
        return {
            "人數": int(len(names)),
            "上班時數": float(wh) if pd.notna(wh) else 0.0,
            "打卡時數": float(ch) if pd.notna(ch) else 0.0,
        }

    stats = {
        "一般人員（不含行政＆幹部）": _calc(df_general),
        "不含幹部（行政有算）": _calc(df_noncadre),
        "全體（所有組別）": _calc(df_all),
        "成箱組": _calc(df_box),
    }

    summary = pd.DataFrame(
        [
            {
                "分類": k,
                "人數": v["人數"],
                "上班時數(O)": v["上班時數"],
                "打卡時數(M)": v["打卡時數"],
            }
            for k, v in stats.items()
        ]
    )

    return {
        "df_base": df_base,
        "df_general": df_general,
        "df_noncadre": df_noncadre,
        "df_all": df_all,
        "df_box": df_box,
        "stats": stats,
        "summary": summary,
        "removed_empty_clockin": removed_c,
        "removed_job": removed_g,
        "total_in": len(df_raw),
        "total_after": len(df_base),
    }


def make_excel_bytes(summary: pd.DataFrame, detail: pd.DataFrame) -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        summary.to_excel(writer, index=False, sheet_name="工時摘要")
        detail.to_excel(writer, index=False, sheet_name="明細")
    return bio.getvalue()


# ----------------------------
# UI
# ----------------------------
card_open("📥 上傳出勤報表")
uploaded = st.file_uploader(
    "請上傳出勤 Excel（.xlsx / .xls）",
    type=["xlsx", "xls", "xlsm"],
    accept_multiple_files=False,
)
card_close()

if not uploaded:
    st.info("請先上傳出勤 Excel 檔。")
    st.stop()

try:
    df_raw = robust_read_excel(uploaded)
    out = build_outputs(df_raw)
except Exception as e:
    st.error(str(e))
    st.stop()

st.caption(
    f"已讀取 {out['total_in']:,} 列；"
    f"排除『上班打卡時間空白』 {out['removed_empty_clockin']:,} 列；"
    f"排除『外倉相關職務』 {out['removed_job']:,} 列；"
    f"剩餘 {out['total_after']:,} 列作為計算基礎。"
)

# =========================
# ✅ KPI 版型：2 欄 × 3 列（你指定的排法）
# =========================
# 第一排：一般人員 vs 不含幹部
c1, c2 = st.columns(2, gap="large")

with c1:
    render_kpis(
        [
            KPI("一般人員｜人數", f"{out['stats']['一般人員（不含行政＆幹部）']['人數']:,}"),
            KPI("一般人員｜上班時數(O)", _fmt2(out["stats"]["一般人員（不含行政＆幹部）"]["上班時數"])),
            KPI("一般人員｜打卡時數(M)", _fmt2(out["stats"]["一般人員（不含行政＆幹部）"]["打卡時數"])),
        ],
        cols=1,
    )

with c2:
    render_kpis(
        [
            KPI("不含幹部｜人數", f"{out['stats']['不含幹部（行政有算）']['人數']:,}"),
            KPI("不含幹部｜上班時數(O)", _fmt2(out["stats"]["不含幹部（行政有算）"]["上班時數"])),
            KPI("不含幹部｜打卡時數(M)", _fmt2(out["stats"]["不含幹部（行政有算）"]["打卡時數"])),
        ],
        cols=1,
    )

st.markdown("")  # 小間距

# 第二排：全體 vs 成箱組
c3, c4 = st.columns(2, gap="large")

with c3:
    render_kpis(
        [
            KPI("全體｜人數", f"{out['stats']['全體（所有組別）']['人數']:,}"),
            KPI("全體｜上班時數(O)", _fmt2(out["stats"]["全體（所有組別）"]["上班時數"])),
            KPI("全體｜打卡時數(M)", _fmt2(out["stats"]["全體（所有組別）"]["打卡時數"])),
        ],
        cols=1,
    )

with c4:
    render_kpis(
        [
            KPI("成箱組｜人數", f"{out['stats']['成箱組']['人數']:,}"),
            KPI("成箱組｜上班時數(O)", _fmt2(out["stats"]["成箱組"]["上班時數"])),
            KPI("成箱組｜打卡時數(M)", _fmt2(out["stats"]["成箱組"]["打卡時數"])),
        ],
        cols=1,
    )

# 摘要
card_open("📌 工時摘要")
df_sum = out["summary"].copy()
df_sum["上班時數(O)"] = df_sum["上班時數(O)"].map(_fmt2)
df_sum["打卡時數(M)"] = df_sum["打卡時數(M)"].map(_fmt2)
st.dataframe(df_sum, use_container_width=True, hide_index=True)
card_close()

# 匯出
card_open("📤 匯出")
scope = st.radio(
    "下載明細範圍",
    options=["一般人員明細", "不含幹部明細", "全體明細", "成箱組明細"],
    horizontal=True,
)

detail_map = {
    "一般人員明細": out["df_general"],
    "不含幹部明細": out["df_noncadre"],
    "全體明細": out["df_all"],
    "成箱組明細": out["df_box"],
}
detail_df = detail_map[scope].copy()

stamp = datetime.now().strftime("%Y%m%d_%H%M")
filename = f"大豐KPI_整體作業工時_{scope}_{stamp}.xlsx"

xlsx_bytes = make_excel_bytes(out["summary"], detail_df)

download_excel_card(
    title="✅ 下載 Excel（含：工時摘要 + 明細）",
    data=xlsx_bytes,
    filename=filename,
)

with st.expander("🔎 明細預覽（前 200 筆）", expanded=False):
    st.dataframe(detail_df.head(200), use_container_width=True)

card_close()
