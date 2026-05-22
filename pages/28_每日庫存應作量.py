# -*- coding: utf-8 -*-
# pages/28_每日庫存應作量.py

import io
import inspect
import pandas as pd
import streamlit as st

# ---- 套用平台風格（有就用，沒有就退回原生）----
try:
    from common_ui import (
        inject_logistics_theme,
        set_page,
        card_open,
        card_close,
        download_excel_card,
    )
    HAS_COMMON_UI = True
except Exception:
    HAS_COMMON_UI = False


# =============================
# helpers
# =============================
def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    new_cols = []
    for c in df.columns:
        s = str(c).replace("\u3000", " ").strip()  # 全形空白 -> 半形 + strip
        new_cols.append(s)
    df.columns = new_cols
    return df


def ensure_order_sku_column(df_order: pd.DataFrame) -> pd.DataFrame:
    """訂單檔：強制對齊『商品』欄位（含：商品 尾巴空白）"""
    df_order = normalize_columns(df_order)
    if "商品" in df_order.columns:
        return df_order

    candidates = ["商品碼", "商品代號", "商品號", "品號", "ITEM", "SKU", "SKU#", "Item", "item"]
    for c in candidates:
        if c in df_order.columns:
            return df_order.rename(columns={c: "商品"})

    raise ValueError(f"訂單檔缺少欄位『商品』。目前欄位：{list(df_order.columns)}")


def format_code(x, length: int) -> str:
    if pd.isna(x) or str(x).strip() == "":
        return ""
    s = str(x).strip().split(".")[0].strip()
    return s.zfill(length)


def _is_fake_xls(raw: bytes) -> bool:
    head = raw[:2048].upper()
    return (b"<HTML" in head) or (b"<TABLE" in head) or (b"PROVIDER" in head)


def _read_fake_xls_text_or_html(raw: bytes) -> pd.DataFrame:
    for enc in ("utf-8-sig", "utf-8", "cp950", "big5"):
        try:
            text = raw.decode(enc, errors="replace")
            break
        except Exception:
            text = None
    if text is None:
        text = raw.decode("utf-8", errors="replace")

    if "<table" in text.lower():
        dfs = pd.read_html(io.StringIO(text))
        if not dfs:
            raise ValueError("偵測為 HTML，但找不到 table")
        return dfs[0]

    # tab -> comma
    try:
        df = pd.read_csv(io.StringIO(text), sep="\t")
        if df.shape[1] >= 2:
            return df
    except Exception:
        pass

    return pd.read_csv(io.StringIO(text), sep=",")


def robust_read_table(uploaded) -> pd.DataFrame:
    name = (uploaded.name or "").lower()
    raw = uploaded.getvalue()

    if name.endswith(".csv"):
        try:
            return pd.read_csv(io.BytesIO(raw), encoding="utf-8-sig")
        except Exception:
            return pd.read_csv(io.BytesIO(raw), encoding="big5", errors="replace")

    if name.endswith((".xlsx", ".xlsm", ".xltx", ".xltm")):
        return pd.read_excel(io.BytesIO(raw), engine="openpyxl")

    if name.endswith(".xls"):
        if _is_fake_xls(raw):
            return _read_fake_xls_text_or_html(raw)
        try:
            import xlrd  # noqa: F401
        except Exception:
            raise ValueError(
                "你上傳的是『真 .xls（舊版 Excel）』，部署環境需安裝 xlrd 才能讀。\n"
                "請在 requirements.txt 加上：xlrd>=2.0.1\n"
                "或先把檔案另存成 .xlsx / .csv 再上傳。"
            )
        return pd.read_excel(io.BytesIO(raw), engine="xlrd")

    return pd.read_excel(io.BytesIO(raw))


def read_master_file(uploaded) -> tuple[pd.DataFrame, pd.DataFrame]:
    raw = uploaded.getvalue()
    name = (uploaded.name or "").lower()

    if name.endswith(".xls"):
        if _is_fake_xls(raw):
            raise ValueError("商品主檔不應是『假 xls』格式，請提供正常 Excel（含分頁）。")
        try:
            import xlrd  # noqa: F401
        except Exception:
            raise ValueError(
                "商品主檔是 .xls，部署環境需安裝 xlrd。\n"
                "請在 requirements.txt 加上：xlrd>=2.0.1\n"
                "或先另存成 .xlsx 再上傳。"
            )
        engine = "xlrd"
    else:
        engine = "openpyxl"

    try:
        df_master = pd.read_excel(io.BytesIO(raw), sheet_name="商品主檔", engine=engine)
        df_weight = pd.read_excel(io.BytesIO(raw), sheet_name="大類加權", engine=engine)
        return normalize_columns(df_master), normalize_columns(df_weight)
    except Exception as e:
        raise ValueError("找不到『商品主檔』或『大類加權』分頁，請檢查 Excel 工作表名稱。") from e


def safe_download_card(label: str, data: bytes, filename: str, mime: str = "text/csv"):
    """相容不同版本 download_excel_card + 最終保底 st.download_button"""
    if HAS_COMMON_UI and "download_excel_card" in globals():
        fn = download_excel_card
        try:
            sig = inspect.signature(fn)
            params = set(sig.parameters.keys())
            kwargs = {}

            for k in ("title", "label", "text"):
                if k in params:
                    kwargs[k] = label
                    break
            for k in ("data", "xlsx_bytes", "bytes_data"):
                if k in params:
                    kwargs[k] = data
                    break
            for k in ("filename", "file_name"):
                if k in params:
                    kwargs[k] = filename
                    break
            if "mime" in params:
                kwargs["mime"] = mime

            if kwargs:
                return fn(**kwargs)
        except Exception:
            pass

        for args in [(label, data, filename), (data, filename), (label, data)]:
            try:
                return fn(*args)
            except Exception:
                continue

    return st.download_button(label, data=data, file_name=filename, mime=mime, use_container_width=True)


def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


def concat_orders(order_files) -> tuple[pd.DataFrame, list[str]]:
    """多檔訂單合併（含 來源檔案 追溯）"""
    msgs = []
    dfs = []
    for f in order_files:
        df = robust_read_table(f)
        df = normalize_columns(df)
        df = ensure_order_sku_column(df)
        df["來源檔案"] = f.name
        dfs.append(df)
        msgs.append(f"已讀取：{f.name}（{len(df):,} 筆）")

    if not dfs:
        raise ValueError("尚未選擇任何訂單檔")

    out = pd.concat(dfs, ignore_index=True, sort=False)
    msgs.append(f"多檔合併完成：{len(dfs)} 檔，共 {len(out):,} 筆")
    return out, msgs


def build_result(df_order: pd.DataFrame, df_master: pd.DataFrame, df_weight: pd.DataFrame):
    msgs: list[str] = []

    df_order = ensure_order_sku_column(df_order)

    # 排除特殊儲位
    exclude_list = ["CGS", "JCPL", "QC99", "PD99", "GX010", "GREAT0001X"]
    if "儲位" in df_order.columns:
        before = len(df_order)
        df_order = df_order.copy()
        df_order["儲位"] = df_order["儲位"].astype(str).str.strip()
        pattern = "|".join(exclude_list)
        df_order = df_order[~df_order["儲位"].str.contains(pattern, case=False, na=False)]
        after = len(df_order)
        msgs.append(f"已排除特殊儲位：{before - after:,} 筆（剩餘 {after:,} 筆）")

    # 必要欄位檢查
    for col in ("商品代號", "大類"):
        if col not in df_master.columns:
            raise ValueError(f"商品主檔缺少欄位『{col}』")
    for col in ("PA", "PARM_VALUE2"):
        if col not in df_weight.columns:
            raise ValueError(f"大類加權分頁缺少欄位『{col}』")

    # 補碼
    df_order = df_order.copy()
    df_master = df_master.copy()
    df_weight = df_weight.copy()

    df_order["商品"] = df_order["商品"].apply(lambda x: format_code(x, 6))
    df_master["商品代號"] = df_master["商品代號"].apply(lambda x: format_code(x, 6))

    df_master["大類"] = df_master["大類"].apply(lambda x: format_code(x, 2))
    df_weight["PA"] = df_weight["PA"].apply(lambda x: format_code(x, 2))

    # Join A：商品 -> 大類
    master_cols = ["商品代號", "大類"]
    if "類別" in df_master.columns:
        master_cols.append("類別")
    df_master_sub = df_master[master_cols].drop_duplicates(subset=["商品代號"])
    step1_df = pd.merge(df_order, df_master_sub, left_on="商品", right_on="商品代號", how="left")

    miss_master_mask = step1_df["大類"].isna() | (step1_df["大類"].astype(str).str.strip() == "")
    miss_master_df = step1_df.loc[miss_master_mask].copy()

    # Join B：大類 -> 加權值
    df_weight_sub = df_weight[["PA", "PARM_VALUE2"]].drop_duplicates(subset=["PA"])
    final_df = pd.merge(step1_df, df_weight_sub, left_on="大類", right_on="PA", how="left")
    final_df = final_df.rename(columns={"PARM_VALUE2": "大類加權值"})

    weight_missing_mask = (
        (~final_df["大類"].isna())
        & (final_df["大類"].astype(str).str.strip() != "")
        & (final_df["大類加權值"].isna() | (final_df["大類加權值"].astype(str).str.strip() == ""))
    )
    miss_weight_df = final_df.loc[weight_missing_mask].copy()

    # 計算加權結果
    qty_col = "計量單位數量"
    weight_col = "大類加權值"
    if qty_col in final_df.columns and weight_col in final_df.columns:
        final_df[qty_col] = pd.to_numeric(final_df[qty_col], errors="coerce").fillna(0)
        final_df[weight_col] = pd.to_numeric(final_df[weight_col], errors="coerce").fillna(0)
        final_df["加權計算結果"] = final_df[weight_col] * final_df[qty_col]
        msgs.append("已完成計算：加權計算結果 = 大類加權值 * 計量單位數量")

    final_df = final_df.drop(columns=["商品代號", "PA"], errors="ignore")
    miss_master_df = miss_master_df.drop(columns=["商品代號"], errors="ignore")
    miss_weight_df = miss_weight_df.drop(columns=["商品代號", "PA"], errors="ignore")

    msgs.append(f"未對應商品主檔（大類空白）：{len(miss_master_df):,} 筆")
    msgs.append(f"未對應大類加權（大類加權值缺失）：{len(miss_weight_df):,} 筆")

    return final_df, msgs, miss_master_df, miss_weight_df


# =============================
# UI
# =============================
def main():
    st.set_page_config(page_title="每日庫存應作量", page_icon="📦", layout="wide")

    # ✅ 你指定：上方只保留「📦 每日庫存應作量」
    if HAS_COMMON_UI:
        inject_logistics_theme()
        # 不要 subtitle（避免出現那串導引文字）
        try:
            set_page("每日庫存應作量", icon="📦")  # 多數 common_ui 版本支援
        except TypeError:
            # 若你的 set_page 沒有 icon 參數，就退回最簡單
            set_page("📦 每日庫存應作量")
    else:
        st.title("📦 每日庫存應作量")

    # 上傳區
    if HAS_COMMON_UI:
        card_open("1) 上傳檔案")

    c1, c2 = st.columns(2)
    with c1:
        order_files = st.file_uploader(
            "訂單資料檔（抓『商品』｜可選擇多檔）",
            type=["csv", "xlsx", "xls", "xlsm"],
            accept_multiple_files=True,
            key="order_multi",
        )
    with c2:
        master_file = st.file_uploader(
            "商品主檔（含：商品主檔 / 大類加權）",
            type=["xlsx", "xls", "xlsm"],
            key="master",
        )

    if HAS_COMMON_UI:
        card_close()

    st.divider()

    run = st.button("✅ 開始處理", type="primary", disabled=not (order_files and master_file))
    if not run:
        return

    try:
        with st.spinner("讀取訂單檔（多檔）..."):
            df_order, read_msgs = concat_orders(order_files)

        with st.spinner("讀取商品主檔..."):
            df_master, df_weight = read_master_file(master_file)

        with st.spinner("處理中（排除 / 補碼 / Join / 計算）..."):
            final_df, msgs, miss_master_df, miss_weight_df = build_result(df_order, df_master, df_weight)

        # ✅ KPI：直向顯示（上下堆疊）
        total_rows = len(final_df)
        uniq_sku = final_df["商品"].nunique() if "商品" in final_df.columns else 0
        sum_weighted = float(final_df["加權計算結果"].sum()) if "加權計算結果" in final_df.columns else 0.0

        st.metric("筆數", f"{total_rows:,}")
        st.metric("商品數(不重複)", f"{uniq_sku:,}")
        st.metric("加權計算結果總和", f"{sum_weighted:,.2f}")

        st.info(" \n".join([f"- {m}" for m in (read_msgs + msgs)]))

        st.subheader("✅ 主結果明細")
        st.dataframe(final_df, use_container_width=True, height=520)

        safe_download_card(
            "✅ 下載 CSV（加權計算結果）",
            to_csv_bytes(final_df),
            "處理完成_加權計算結果.csv",
            mime="text/csv",
        )

        st.divider()

        with st.expander("⚠️ 比對不到『大類加權』的明細（大類存在但加權值缺失）", expanded=(len(miss_weight_df) > 0)):
            if len(miss_weight_df) == 0:
                st.success("沒有比對不到的大類 ✅")
            else:
                cat_summary = (
                    miss_weight_df.assign(大類=miss_weight_df["大類"].astype(str).str.strip())
                    .groupby("大類", dropna=False)
                    .size()
                    .reset_index(name="筆數")
                    .sort_values("筆數", ascending=False)
                )
                st.dataframe(cat_summary, use_container_width=True, height=240)
                st.dataframe(miss_weight_df, use_container_width=True, height=520)
                safe_download_card(
                    "⬇️ 下載 CSV（比對不到大類加權-明細）",
                    to_csv_bytes(miss_weight_df),
                    "未對應大類加權_明細.csv",
                    mime="text/csv",
                )

        with st.expander("⚠️ 商品找不到『商品主檔』明細（大類空白）", expanded=False):
            if len(miss_master_df) == 0:
                st.success("全部商品皆能對應商品主檔 ✅")
            else:
                st.dataframe(miss_master_df, use_container_width=True, height=520)
                safe_download_card(
                    "⬇️ 下載 CSV（商品未對應主檔-明細）",
                    to_csv_bytes(miss_master_df),
                    "未對應商品主檔_明細.csv",
                    mime="text/csv",
                )

        st.success("完成 ✅")

    except Exception as e:
        st.error(f"執行中發生問題：{e}")


if __name__ == "__main__":
    main()
