# pages/21_進貨課 - 上架量體.py
# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import re
from io import BytesIO
from typing import Dict, Optional, List, Tuple

import pandas as pd
import streamlit as st

from common_ui import inject_logistics_theme, set_page, card_open, card_close

pd.options.display.max_columns = 200

# =========================
# 設定（沿用你原腳本）
# =========================
TO_EXCLUDE_KEYWORDS = ["CGS", "JCPL", "QC99", "GREAT0001X", "GX010", "PD99"]
TO_EXCLUDE_PATTERN = re.compile("|".join(re.escape(k) for k in TO_EXCLUDE_KEYWORDS), flags=re.IGNORECASE)

ITEM_CANDIDATES = [
    "ITEM", "Item", "item", "商品", "品號", "品項", "商品代號", "貨號", "料號",
    "itemcode", "ItemCode", "ITEMCODE",
]

LOC_KEY_CANDIDATES = ["儲位", "儲位代碼", "到", "儲位編號", "Location", "LOC", "loc"]
LOC_TYPE_COL = "儲位類型"


# =========================
# 讀檔（部署版：bytes）
# =========================
def _read_csv_auto(file_bytes: bytes) -> pd.DataFrame:
    last_err = None
    for enc in ("utf-8-sig", "utf-8", "cp950", "big5"):
        try:
            return pd.read_csv(BytesIO(file_bytes), encoding=enc, low_memory=False)
        except Exception as e:
            last_err = e
    raise RuntimeError(f"CSV/TXT 讀取失敗（已嘗試 utf-8-sig/utf-8/cp950/big5）：{last_err}")


def _read_excel_sheets(file_bytes: bytes, ext: str) -> Dict[str, pd.DataFrame]:
    bio = BytesIO(file_bytes)

    if ext in (".xlsx", ".xlsm", ".xltx", ".xltm"):
        try:
            return pd.read_excel(bio, sheet_name=None, engine="openpyxl")
        except Exception:
            bio.seek(0)
            return pd.read_excel(bio, sheet_name=None)

    if ext == ".xls":
        # 若環境沒 xlrd，會提示你先另存 xlsx
        try:
            return pd.read_excel(bio, sheet_name=None, engine="xlrd")
        except Exception as e:
            raise RuntimeError("目前環境可能未安裝 xlrd，.xls 無法讀取；請先另存為 .xlsx 再上傳。") from e

    if ext == ".xlsb":
        try:
            return pd.read_excel(bio, sheet_name=None, engine="pyxlsb")
        except Exception as e:
            raise RuntimeError("目前環境可能未安裝 pyxlsb，.xlsb 無法讀取；請先另存為 .xlsx 再上傳。") from e

    # fallback
    bio.seek(0)
    return pd.read_excel(bio, sheet_name=None)


def read_any_table_from_upload(uploaded) -> Dict[str, pd.DataFrame]:
    name = uploaded.name
    _, ext = os.path.splitext(name)
    ext = ext.lower()
    b = uploaded.getvalue()

    if ext in (".csv", ".txt"):
        return {"CSV": _read_csv_auto(b)}
    return _read_excel_sheets(b, ext)


# =========================
# 共用工具（沿用你原腳本）
# =========================
def normalize_to_qc(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip().str.upper()
    return s.eq("QC")


def to_not_excluded_mask(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip()
    hit = s.str.contains(TO_EXCLUDE_PATTERN, na=False)
    return ~hit


def rename_item_column(df: pd.DataFrame) -> pd.DataFrame:
    cols = list(df.columns)
    if "ITEM" in cols:
        return df
    for cand in ITEM_CANDIDATES:
        if cand in cols and cand != "ITEM":
            return df.rename(columns={cand: "ITEM"})
    return df


def detect_col(df: pd.DataFrame, candidates) -> Optional[str]:
    for c in candidates:
        if c in df.columns:
            return c
    return None


def classify_high_low(storage_type: str) -> str:
    if pd.isna(storage_type):
        return "無法對應"
    s = str(storage_type)
    if "高空" in s:
        return "高空"
    if any(k in s for k in ["低空", "輕型", "落地", "重型", "GM"]):
        return "低空"
    return "未知"


def process_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    # 1) 由=QC（若無『由』→ 全部不保留）
    mask_qc = normalize_to_qc(df["由"]) if "由" in df.columns else pd.Series(False, index=df.index)
    # 2) 到 不含排除關鍵字（若無『到』→ 視為通過）
    mask_to_ok = to_not_excluded_mask(df["到"]) if "到" in df.columns else pd.Series(True, index=df.index)
    return df[mask_qc & mask_to_ok].copy()


def attach_storage_type(df: pd.DataFrame, loc_map: pd.DataFrame, loc_key_map: str) -> pd.DataFrame:
    if "到" not in df.columns:
        df["儲位類型"] = pd.NA
        df["高低空"] = "無法對應"
        return df

    map_df = loc_map.copy()
    df["_到_key_"] = df["到"].astype(str).str.strip()
    map_df["_loc_key_"] = map_df[loc_key_map].astype(str).str.strip()

    if LOC_TYPE_COL not in map_df.columns:
        raise RuntimeError(f"儲位明細缺少必要欄位「{LOC_TYPE_COL}」。")

    map_df = map_df[["_loc_key_", LOC_TYPE_COL]].drop_duplicates("_loc_key_")
    out = df.merge(map_df, how="left", left_on="_到_key_", right_on="_loc_key_")
    out = out.drop(columns=["_到_key_", "_loc_key_"])
    out["高低空"] = out[LOC_TYPE_COL].apply(classify_high_low)
    return out


def _safe_sheet_name(name: str) -> str:
    n = str(name).replace("/", "_").replace("\\", "_").strip()
    return (n[:31] if len(n) > 31 else n) or "Sheet"


def _kpi_text(title: str, value: int):
    st.markdown(f"**{title}**")
    st.markdown(
        f"<div style='font-size:24px; font-weight:900; line-height:1.1; margin-top:2px; margin-bottom:12px;'>{value:,}</div>",
        unsafe_allow_html=True,
    )


def build_output_excel_bytes(
    processed_by_sheet: Dict[str, pd.DataFrame],
    summary_df: pd.DataFrame,
    type_dist_df: Optional[pd.DataFrame],
) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for sn, df in processed_by_sheet.items():
            df.to_excel(writer, sheet_name=_safe_sheet_name(sn), index=False)

        summary_df.to_excel(writer, sheet_name="彙總", index=False)

        if type_dist_df is not None and not type_dist_df.empty:
            type_dist_df.to_excel(writer, sheet_name="儲位類型分佈", index=False)

    out.seek(0)
    return out.read()


# =========================
# UI
# =========================
inject_logistics_theme()
set_page("進貨課｜上架量體", icon="📦", subtitle="由=QC｜到排除關鍵字｜對應儲位類型｜高低空統計｜輸出Excel")

card_open("📦 上架量體（由=QC + 儲位高低空）")

main_up = st.file_uploader(
    "上傳主檔（Excel / CSV）",
    type=["xlsx", "xlsm", "xltx", "xltm", "xls", "xlsb", "csv", "txt"],
    accept_multiple_files=False,
)

storage_up = st.file_uploader(
    "上傳儲位明細（需含：儲位類型 + 儲位鍵欄位）",
    type=["xlsx", "xlsm", "xltx", "xltm", "xls", "xlsb", "csv", "txt"],
    accept_multiple_files=False,
)

st.markdown("---")

if (main_up is None) or (storage_up is None):
    st.info("請先上傳「主檔」與「儲位明細」。")
    card_close()
    st.stop()

run = st.button("開始產出", type="primary")
if not run:
    card_close()
    st.stop()

# 讀檔
try:
    main_sheets = read_any_table_from_upload(main_up)
    sto_sheets = read_any_table_from_upload(storage_up)
except Exception as e:
    st.error(f"讀檔失敗：{e}")
    card_close()
    st.stop()

# 儲位明細：取第一張表
sto_first = list(sto_sheets.keys())[0]
sto_df = sto_sheets[sto_first].copy()
sto_df.columns = [str(c).strip() for c in sto_df.columns]

sto_loc_col = detect_col(sto_df, LOC_KEY_CANDIDATES)
if sto_loc_col is None:
    st.error(f"儲位明細找不到儲位鍵欄位（候選：{', '.join(LOC_KEY_CANDIDATES)}）。")
    card_close()
    st.stop()

if LOC_TYPE_COL not in sto_df.columns:
    st.error(f"儲位明細缺少欄位「{LOC_TYPE_COL}」。")
    card_close()
    st.stop()

# 處理
processed_by_sheet: Dict[str, pd.DataFrame] = {}
summary_rows: List[dict] = []
totals = {"ITEM": 0, "高空": 0, "低空": 0, "未知": 0, "無法對應": 0}

all_concat: List[pd.DataFrame] = []

for sn, df in main_sheets.items():
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = rename_item_column(df)

    kept = process_dataframe(df)
    kept2 = attach_storage_type(kept, sto_df, sto_loc_col)

    # 計數（ITEM=筆數）
    cnt = int(kept2.shape[0])
    c_high = int((kept2["高低空"] == "高空").sum()) if "高低空" in kept2.columns else 0
    c_low = int((kept2["高低空"] == "低空").sum()) if "高低空" in kept2.columns else 0
    c_unknown = int((kept2["高低空"] == "未知").sum()) if "高低空" in kept2.columns else 0
    c_nomap = int((kept2["高低空"] == "無法對應").sum()) if "高低空" in kept2.columns else 0

    processed_by_sheet[sn] = kept2
    all_concat.append(kept2.assign(_Sheet=str(sn)))

    totals["高空"] += c_high
    totals["低空"] += c_low
    totals["ITEM"] += cnt
    totals["未知"] += c_unknown
    totals["無法對應"] += c_nomap

    summary_rows.append(
        {"Sheet": str(sn), "ITEM": cnt, "高空": c_high, "低空": c_low, "未知": c_unknown, "無法對應": c_nomap}
    )

summary_rows.append(
    {"Sheet": "ALL", "ITEM": totals["ITEM"], "高空": totals["高空"], "低空": totals["低空"], "未知": totals["未知"], "無法對應": totals["無法對應"]}
)
summary_df = pd.DataFrame(summary_rows)

# 儲位類型分佈（整體）
type_dist_df = None
all_df = pd.concat(all_concat, ignore_index=True) if all_concat else pd.DataFrame()
if not all_df.empty and LOC_TYPE_COL in all_df.columns:
    type_dist_df = (
        all_df.groupby(LOC_TYPE_COL, dropna=False)
        .size()
        .reset_index(name="ITEM")
        .sort_values("ITEM", ascending=False)
        .reset_index(drop=True)
    )

# 顯示 KPI（純文字、直向）
st.markdown("### 上架量體")
_kpi_text("低空", totals["低空"])
_kpi_text("高空", totals["高空"])
_kpi_text("ITEM（由=QC 且 到通過排除）", totals["ITEM"])
_kpi_text("未知", totals["未知"])
_kpi_text("無法對應", totals["無法對應"])

st.markdown("### 彙總（逐表）")
st.dataframe(summary_df, use_container_width=True, hide_index=True)

if type_dist_df is not None and not type_dist_df.empty:
    st.markdown("### 儲位類型分佈（整體）")
    st.dataframe(type_dist_df, use_container_width=True, hide_index=True)

with st.expander("🔎 預覽：過濾後明細（每張表前 200 筆）", expanded=False):
    for sn, df in processed_by_sheet.items():
        st.markdown(f"**{_safe_sheet_name(sn)}**（{len(df):,} 筆）")
        st.dataframe(df.head(200), use_container_width=True, hide_index=True)

# 產出 Excel
try:
    base, _ = os.path.splitext(main_up.name)
    out_bytes = build_output_excel_bytes(processed_by_sheet, summary_df, type_dist_df)
    out_name = f"{base}_ITEM_高低空.xlsx"
except Exception as e:
    st.error(f"寫檔失敗：{e}")
    card_close()
    st.stop()

st.download_button(
    "⬇️ 下載輸出 Excel",
    data=out_bytes,
    file_name=out_name,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

card_close()
