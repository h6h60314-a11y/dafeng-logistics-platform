# pages/23_採品門市差異量.py
# -*- coding: utf-8 -*-
import pandas as pd
import streamlit as st
from io import BytesIO, StringIO

from common_ui import inject_logistics_theme, set_page, card_open, card_close


# ----------------------------
# helpers
# ----------------------------
REQUIRED_COLS = [
    "提供日期",
    "驗收日",
    "採購單號",
    "供應商代號",
    "廠商名",
    "商品碼",
    "數量",
    "門市代碼",
    "門市名",
    "未配出原因",
    "備註",
]


def _as_text(x):
    if x is None:
        return ""
    try:
        if pd.isna(x):
            return ""
    except Exception:
        pass
    return str(x)


def _read_excel(uploaded_file, sheet_name=0) -> pd.DataFrame:
    return pd.read_excel(uploaded_file, sheet_name=sheet_name, engine="openpyxl")


def _read_excel_all_sheets(uploaded_file) -> dict:
    return pd.read_excel(uploaded_file, sheet_name=None, engine="openpyxl")


def _ensure_cols(df: pd.DataFrame, cols: list) -> pd.DataFrame:
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    front = [c for c in cols if c in df.columns]
    tail = [c for c in df.columns if c not in front]
    return df[front + tail]


def _build_output_bytes(sheets: dict) -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        for name, df in sheets.items():
            safe_name = str(name)[:31]
            df.to_excel(writer, sheet_name=safe_name, index=False)
    bio.seek(0)
    return bio.getvalue()


def _read_pasted_table(text: str) -> pd.DataFrame:
    """
    支援從 Excel 複製貼上：
    - 通常是 TAB 分隔（\t）
    - 若是 CSV 也可（,）
    """
    raw = (text or "").strip("\n").strip()
    if not raw:
        raise ValueError("貼上的內容是空的。請從 Excel 複製整段（含表頭）再貼上。")

    # 先猜 TAB（Excel 最常見）
    try:
        df = pd.read_csv(StringIO(raw), sep="\t", dtype=str)
        if df.shape[1] <= 1:
            raise ValueError("not tab")
        return df
    except Exception:
        pass

    # 再猜 CSV
    try:
        df = pd.read_csv(StringIO(raw), sep=",", dtype=str)
        if df.shape[1] <= 1:
            raise ValueError("not csv")
        return df
    except Exception:
        pass

    # 最後：嘗試用任意空白（很少見）
    df = pd.read_csv(StringIO(raw), sep=r"\s+", dtype=str)
    if df.shape[1] <= 1:
        raise ValueError("無法解析貼上內容：請確認是『含表頭』且有分隔符（Excel 複製通常為 TAB）。")
    return df


# ----------------------------
# page
# ----------------------------
st.set_page_config(page_title="大豐物流｜採品門市差異量", page_icon="📄", layout="wide")
inject_logistics_theme()
set_page("📄 採品門市差異量（依未配出原因回填分頁）", "出貨課｜採品／門市差異彙整")

card_open("操作說明")
st.markdown(
    """
- 準備 **2 個來源**：  
  1) **採品明細**：請直接在平台用「複製貼上」（從 Excel 複製整塊資料 *含表頭*）  
  2) **採品門市差異量**：上傳多分頁 Excel（分頁名稱 = `未配出原因`）
- 系統會把「採品明細」逐筆依 `未配出原因` 追加到對應分頁。
- 僅當 `未配出原因` **有對應分頁名稱** 時才會寫入；找不到分頁的會列在「未對應清單」。
"""
)
card_close()

st.divider()

# ----------- 採品明細輸入方式 -----------
card_open("① 採品明細來源")
mode = st.radio(
    "選擇輸入方式",
    ["複製貼上（推薦）", "上傳 Excel（備用）"],
    horizontal=True,
)

df_detail = None

if mode == "複製貼上（推薦）":
    pasted = st.text_area(
        "把採品明細從 Excel 複製後貼在這裡（請包含表頭）",
        height=220,
        placeholder="在 Excel 選取含表頭的整段資料 → Ctrl+C → 這裡 Ctrl+V",
    )
    parse_btn = st.button("解析貼上內容", type="primary", use_container_width=False)

    if parse_btn:
        try:
            df_detail = _read_pasted_table(pasted)
            st.session_state["df_detail_pasted"] = df_detail
            st.success(f"解析成功：{df_detail.shape[0]:,} 筆 × {df_detail.shape[1]} 欄")
        except Exception as e:
            st.error(f"解析失敗：{e}")

    # 若已解析過，沿用 session_state
    if "df_detail_pasted" in st.session_state and df_detail is None:
        df_detail = st.session_state["df_detail_pasted"]

else:
    f_detail = st.file_uploader("上傳：採品明細（.xlsx）", type=["xlsx"], accept_multiple_files=False)
    if f_detail:
        try:
            df_detail = _read_excel(f_detail, sheet_name=0)
            st.success(f"讀取成功：{df_detail.shape[0]:,} 筆 × {df_detail.shape[1]} 欄")
        except Exception as e:
            st.error(f"採品明細讀取失敗：{e}")

card_close()

st.divider()

# ----------- 上傳差異量活頁簿 -----------
card_open("② 採品門市差異量（多分頁 Excel）")
f_book = st.file_uploader(
    "上傳：採品門市差異量（多分頁 .xlsx）",
    type=["xlsx"],
    accept_multiple_files=False,
)
card_close()

st.divider()

# 必要輸入檢查
if df_detail is None:
    st.info("請先完成『① 採品明細』貼上解析或上傳。")
    st.stop()

if not f_book:
    st.info("請上傳『② 採品門市差異量（多分頁 Excel）』。")
    st.stop()

# 讀取多分頁
try:
    sheets = _read_excel_all_sheets(f_book)
except Exception as e:
    st.error(f"採品門市差異量（多分頁）讀取失敗：{e}")
    st.stop()

# 檢查必要欄位（至少要有 未配出原因）
if "未配出原因" not in df_detail.columns:
    st.error("採品明細缺少必要欄位：未配出原因（請確認貼上/上傳資料的表頭名稱）")
    st.stop()

# 若採品明細沒有「備註」，也先補一個空欄
if "備註" not in df_detail.columns:
    df_detail["備註"] = ""

# 統一欄位
df_detail = _ensure_cols(df_detail.copy(), REQUIRED_COLS)

# 各分頁補齊欄位
for k in list(sheets.keys()):
    try:
        sheets[k] = _ensure_cols(sheets[k].copy(), REQUIRED_COLS)
    except Exception:
        sheets[k] = pd.DataFrame(columns=REQUIRED_COLS)

# 主邏輯：依未配出原因回填
matched = 0
skipped = 0
missing_reasons = []

for _, row in df_detail.iterrows():
    reason = _as_text(row.get("未配出原因")).strip()
    if not reason:
        skipped += 1
        continue

    if reason in sheets:
        new_row = pd.DataFrame([{c: row.get(c, "") for c in REQUIRED_COLS}])
        sheets[reason] = pd.concat([sheets[reason], new_row], ignore_index=True)
        matched += 1
    else:
        missing_reasons.append(reason)
        skipped += 1

# 統計展示
card_open("處理結果")
c1, c2, c3 = st.columns(3)
c1.metric("寫入筆數", f"{matched:,}")
c2.metric("略過筆數", f"{skipped:,}")
c3.metric("分頁總數", f"{len(sheets):,}")
card_close()

if missing_reasons:
    uniq_missing = sorted(set([x for x in missing_reasons if x]))
    with st.expander(f"未對應分頁的 未配出原因（{len(uniq_missing)} 種）", expanded=False):
        st.write(uniq_missing)

# 下載
out_bytes = _build_output_bytes(sheets)
out_name = "更新後的採品門市差異量.xlsx"

st.download_button(
    label="⬇️ 下載：更新後的採品門市差異量.xlsx",
    data=out_bytes,
    file_name=out_name,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

# 預覽
with st.expander("預覽：採品明細（前 200 筆）", expanded=False):
    st.dataframe(df_detail.head(200), use_container_width=True)

with st.expander("預覽：分頁內容（選一張）", expanded=False):
    sheet_names = list(sheets.keys())
    pick = st.selectbox("分頁", sheet_names, index=0 if sheet_names else None)
    if pick:
        st.dataframe(sheets[pick].head(200), use_container_width=True)
