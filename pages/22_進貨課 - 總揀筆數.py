# pages/22_進貨課 - 總揀筆數.py
# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import re
from io import BytesIO
from typing import List, Tuple

import numpy as np
import pandas as pd
import streamlit as st

from common_ui import inject_logistics_theme, set_page, card_open, card_close

pd.options.display.max_columns = 200

# 需要排除的儲位關鍵字（子字串比對，不分大小寫）
EXCLUDE_SUBSTRINGS = ["CGS", "JCPL", "QC99", "GREAT0001X", "GX010", "PD99"]
EXCLUDE_PATTERN = re.compile("|".join(map(re.escape, EXCLUDE_SUBSTRINGS)), re.IGNORECASE)

# 樞紐#1 的列階層（中繼用，動態擇用）
PIVOT1_BASE_ROWS = ["儲位類型", "儲位", "商品"]  # 一定會用
PIVOT1_OPTIONAL = ["揀貨批次號"]               # 若存在才用

# 供樞紐#1 內部加總（若欄位存在才會合計）
NUM_COL_CANDIDATES = ["數量", "件數", "出貨量", "配貨量", "應揀量", "RF揀貨量", "差異量"]

# 若來源沒有商品欄位，補虛擬欄位
PRODUCT_FALLBACK_COL = "商品"


def normalize_loc(s):
    if pd.isna(s):
        return s
    return str(s).strip().upper()


def unit_mask_equal_2(series: pd.Series) -> pd.Series:
    """成箱：=2（字面 '2' 或數值 2/2.0）"""
    s = series.astype(str).str.strip()
    mask_str = s.eq("2")
    s_num = pd.to_numeric(s, errors="coerce")
    mask_num = np.isfinite(s_num) & np.isclose(s_num, 2.0, rtol=0, atol=1e-9)
    return mask_str | mask_num


def unit_mask_contains_3_or_6(series: pd.Series) -> pd.Series:
    """零散：字串含 3 或 6（含全形 ３／６），任意位置"""
    s = series.astype(str)
    pat = re.compile(r"[3３]|[6６]")
    return s.str.contains(pat, na=False)


def _read_csv_auto(file_bytes: bytes) -> pd.DataFrame:
    last_err = None
    for enc in ("utf-8-sig", "utf-8", "cp950", "big5"):
        try:
            return pd.read_csv(BytesIO(file_bytes), encoding=enc, low_memory=False, dtype=str)
        except Exception as e:
            last_err = e
    raise RuntimeError(f"CSV/TXT 讀取失敗（已嘗試 utf-8-sig/utf-8/cp950/big5）：{last_err}")


def read_excel_or_csv(uploaded) -> pd.DataFrame:
    """讀單表（與你原本 read_excel 行為一致），支援 Excel/CSV/TXT"""
    name = uploaded.name
    _, ext = os.path.splitext(name)
    ext = ext.lower()
    b = uploaded.getvalue()

    if ext in (".csv", ".txt"):
        return _read_csv_auto(b)

    bio = BytesIO(b)
    if ext in (".xlsx", ".xlsm", ".xltx", ".xltm"):
        try:
            return pd.read_excel(bio, dtype=str, engine="openpyxl")
        except Exception:
            bio.seek(0)
            return pd.read_excel(bio, dtype=str)

    if ext == ".xls":
        try:
            return pd.read_excel(bio, dtype=str, engine="xlrd")
        except Exception as e:
            raise RuntimeError("目前環境可能未安裝 xlrd，.xls 無法讀取；請先另存為 .xlsx 再上傳。") from e

    if ext == ".xlsb":
        try:
            return pd.read_excel(bio, dtype=str, engine="pyxlsb")
        except Exception as e:
            raise RuntimeError("目前環境可能未安裝 pyxlsb，.xlsb 無法讀取；請先另存為 .xlsx 再上傳。") from e

    bio.seek(0)
    return pd.read_excel(bio, dtype=str)


def build_pivot2(df_source: pd.DataFrame) -> Tuple[pd.DataFrame, List[str]]:
    """
    回傳 (#2 DataFrame, 實際分組鍵清單)
    - 若有「揀貨批次號」：["儲位類型","揀貨批次號"]
    - 若無「揀貨批次號」：["儲位類型"]
    """
    df_tmp = df_source.copy()

    # 供中繼合計使用（不影響筆數）
    for c in [c for c in NUM_COL_CANDIDATES if c in df_tmp.columns]:
        df_tmp[c] = pd.to_numeric(df_tmp[c], errors="coerce")

    gb1 = [k for k in (PIVOT1_OPTIONAL + PIVOT1_BASE_ROWS) if k in df_tmp.columns]
    if ("儲位類型" not in gb1) or ("儲位" not in gb1):
        raise ValueError("樞紐所需欄位不足，至少要有：儲位、儲位類型。")

    # pivot1：中繼（按 可選批次號 + 儲位類型 + 儲位 + 商品）
    pivot1 = (
        df_tmp.groupby(gb1, dropna=False)
        .size()
        .reset_index(name="筆數")
    )

    group_keys = ["儲位類型"] + (["揀貨批次號"] if "揀貨批次號" in pivot1.columns else [])

    # pivot2：輸出（按 儲位類型(+批次號) 統計 儲位筆數）
    pivot2 = (
        pivot1.groupby(group_keys, dropna=False)["儲位"]
        .count()
        .reset_index(name="儲位_筆數")
        .sort_values(group_keys, kind="mergesort")
        .reset_index(drop=True)
    )
    pivot2["儲位_筆數"] = pivot2["儲位_筆數"].fillna(0).astype(int)
    return pivot2, group_keys


def process_subset(df_raw: pd.DataFrame, df_map: pd.DataFrame, subset_tag: str, mask: pd.Series):
    df_work = df_raw.loc[mask].copy()
    if df_work.empty:
        return subset_tag, None, 0, None

    # 排除儲位
    df_work = df_work[~df_work["儲位"].astype(str).str.contains(EXCLUDE_PATTERN, na=False)].copy()
    if df_work.empty:
        return subset_tag, None, 0, None

    # 回填儲位類型
    df_work["儲位_norm"] = df_work["儲位"].map(normalize_loc)
    map_first = (
        df_map.assign(儲位_norm=df_map["儲位"].map(normalize_loc))
        .sort_values(["儲位_norm"])
        .drop_duplicates(subset=["儲位_norm"], keep="first")[["儲位_norm", "儲位類型"]]
    )
    df_out = df_work.merge(map_first, on="儲位_norm", how="left").drop(columns=["儲位_norm"])

    pivot2, group_keys = build_pivot2(df_out)
    total_count = int(pivot2["儲位_筆數"].sum(skipna=True))
    return subset_tag, pivot2, total_count, group_keys


def build_single_sheet_excel_bytes(df_type_total: pd.DataFrame, df_detail_all: pd.DataFrame, df_summary: pd.DataFrame) -> bytes:
    """
    單一工作表（依需求順序）：
      1) 總揀筆數（最上方）
      2) 明細表（合併）
      3) 彙總總表
    """
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        sheet = "結果"
        r = 0

        # 1) 總揀筆數（依儲位類型加總）
        pd.DataFrame({"": ["總揀筆數"]}).to_excel(
            writer, sheet_name=sheet, index=False, header=False, startrow=r, startcol=0
        )
        r += 1
        df_type_total.to_excel(writer, sheet_name=sheet, index=False, startrow=r, startcol=0)
        r += len(df_type_total) + 2

        # 2) 明細表（合併）
        pd.DataFrame({"": ["明細表（合併）"]}).to_excel(
            writer, sheet_name=sheet, index=False, header=False, startrow=r, startcol=0
        )
        r += 1
        df_detail_all.to_excel(writer, sheet_name=sheet, index=False, startrow=r, startcol=0)
        r += len(df_detail_all) + 2

        # 3) 彙總總表
        pd.DataFrame({"": ["彙總總表"]}).to_excel(
            writer, sheet_name=sheet, index=False, header=False, startrow=r, startcol=0
        )
        r += 1
        df_summary.to_excel(writer, sheet_name=sheet, index=False, startrow=r, startcol=0)

    out.seek(0)
    return out.read()


def _label_type_as_pick(type_name: str) -> str:
    t = str(type_name).strip()
    if t == "低空":
        return "低空總揀筆數"
    if t == "高空":
        return "高空總揀筆數"
    if t.upper() == "GM":
        return "GM總揀筆數"
    return f"{t}總揀筆數"


def show_type_totals_as_text(df_type_total: pd.DataFrame):
    """✅ 不用表格：純文字直向顯示（標題與名稱依需求）"""
    st.markdown("### 總揀筆數")
    if df_type_total is None or df_type_total.empty:
        st.caption("（無資料）")
        return

    for _, r in df_type_total.iterrows():
        t = r.get("儲位類型", "")
        v = r.get("總揀筆數", 0)
        st.markdown(f"**{_label_type_as_pick(t)}**")
        st.markdown(
            f"<div style='font-size:28px; font-weight:900; line-height:1.1; margin-top:2px; margin-bottom:14px;'>{int(v):,}</div>",
            unsafe_allow_html=True,
        )


# =========================
# UI
# =========================
inject_logistics_theme()
set_page("進貨課｜總揀筆數", icon="🎯", subtitle="多檔批次｜成箱/零散（或ALL）｜排除儲位｜回填儲位類型｜單頁Excel輸出")

card_open("🎯 總揀筆數（單頁輸出）")

batch_files = st.file_uploader(
    "上傳【批次明細】（可多檔；至少含欄位：儲位；有『計量單位』更佳）",
    type=["xlsx", "xlsm", "xltx", "xltm", "xls", "xlsb", "csv", "txt"],
    accept_multiple_files=True,
)

map_file = st.file_uploader(
    "上傳【儲位棚別明細】（需含欄位：儲位、儲位類型）",
    type=["xlsx", "xlsm", "xltx", "xltm", "xls", "xlsb", "csv", "txt"],
    accept_multiple_files=False,
)

st.markdown("---")

if (not batch_files) or (map_file is None):
    st.info("請先上傳「批次明細（可多檔）」與「儲位棚別明細」。")
    card_close()
    st.stop()

run = st.button("開始產出", type="primary")
if not run:
    card_close()
    st.stop()

# 讀取 map
try:
    df_map = read_excel_or_csv(map_file)
    df_map.columns = [str(c).strip() for c in df_map.columns]
except Exception as e:
    st.error(f"讀取『儲位棚別明細』失敗：{e}")
    card_close()
    st.stop()

for c in ["儲位", "儲位類型"]:
    if c not in df_map.columns:
        st.error(f"儲位棚別明細缺少欄位：{c}")
        card_close()
        st.stop()

summary_rows: List[dict] = []
detail_frames: List[pd.DataFrame] = []
ok, fail = 0, 0

for up in batch_files:
    base = os.path.basename(up.name)
    name_noext = os.path.splitext(base)[0]

    try:
        df_raw = read_excel_or_csv(up)
        df_raw.columns = [str(c).strip() for c in df_raw.columns]
    except Exception:
        summary_rows.append({"來源檔名": name_noext, "子集": "讀檔失敗", "分組鍵": "無", "加總筆數": 0, "資料筆數": 0})
        fail += 1
        continue

    if "儲位" not in df_raw.columns:
        summary_rows.append({"來源檔名": name_noext, "子集": "缺欄位", "分組鍵": "無", "加總筆數": 0, "資料筆數": 0})
        fail += 1
        continue

    if "商品" not in df_raw.columns:
        df_raw[PRODUCT_FALLBACK_COL] = 1
        df_raw = df_raw.rename(columns={PRODUCT_FALLBACK_COL: "商品"})

    has_unit_col = "計量單位" in df_raw.columns
    masks = (
        [("成箱", unit_mask_equal_2(df_raw["計量單位"])),
         ("零散", unit_mask_contains_3_or_6(df_raw["計量單位"]))]
        if has_unit_col
        else [("ALL", pd.Series([True] * len(df_raw), index=df_raw.index))]
    )

    any_ok = False
    for tag, mask in masks:
        tag, pivot2, total_count, group_keys = process_subset(df_raw, df_map, tag, mask)

        if pivot2 is None:
            summary_rows.append({"來源檔名": name_noext, "子集": tag, "分組鍵": "無", "加總筆數": 0, "資料筆數": 0})
            continue

        grp_desc = " × ".join(group_keys)

        df_detail = pivot2.copy()
        df_detail.insert(0, "來源檔名", name_noext)
        df_detail.insert(1, "子集", tag)
        detail_frames.append(df_detail)

        summary_rows.append(
            {"來源檔名": name_noext, "子集": tag, "分組鍵": grp_desc, "加總筆數": int(total_count), "資料筆數": int(len(pivot2))}
        )
        any_ok = True

    if any_ok:
        ok += 1
    else:
        fail += 1

df_summary = pd.DataFrame(summary_rows) if summary_rows else pd.DataFrame(columns=["來源檔名", "子集", "分組鍵", "加總筆數", "資料筆數"])
if not df_summary.empty:
    df_summary = df_summary.sort_values(["來源檔名", "子集"], kind="mergesort").reset_index(drop=True)

if detail_frames:
    df_detail_all = pd.concat(detail_frames, ignore_index=True)
    base_cols = ["來源檔名", "子集", "儲位類型"]
    cols = base_cols + (["揀貨批次號"] if "揀貨批次號" in df_detail_all.columns else []) + ["儲位_筆數"]
    others = [c for c in df_detail_all.columns if c not in cols]
    df_detail_all = df_detail_all[cols + others]
else:
    df_detail_all = pd.DataFrame(columns=["來源檔名", "子集", "儲位類型", "儲位_筆數"])

# ✅ 依儲位類型加總「總揀筆數」（儲位_筆數 sum）
if (not df_detail_all.empty) and ("儲位類型" in df_detail_all.columns) and ("儲位_筆數" in df_detail_all.columns):
    df_type_total = (
        df_detail_all.groupby("儲位類型", dropna=False)["儲位_筆數"]
        .sum()
        .reset_index(name="總揀筆數")
        .sort_values("總揀筆數", ascending=False, kind="mergesort")
        .reset_index(drop=True)
    )
else:
    df_type_total = pd.DataFrame(columns=["儲位類型", "總揀筆數"])

# ✅ 顯示（不要表格）
show_type_totals_as_text(df_type_total)

# ✅ 你要的順序：明細表（合併）在彙總總表上
st.markdown("### 明細表（合併）")
st.dataframe(df_detail_all, use_container_width=True, hide_index=True)

st.markdown("### 彙總總表")
st.dataframe(df_summary, use_container_width=True, hide_index=True)

st.caption(f"成功：{ok} 檔；失敗：{fail} 檔")

# 下載（Excel：同一張工作表，順序同畫面）
try:
    out_bytes = build_single_sheet_excel_bytes(df_type_total, df_detail_all, df_summary)
    out_name = "批次_總揀筆數_單頁輸出.xlsx"
except Exception as e:
    st.error(f"輸出失敗：{e}")
    card_close()
    st.stop()

st.download_button(
    "⬇️ 下載輸出 Excel（單一工作表）",
    data=out_bytes,
    file_name=out_name,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

card_close()
