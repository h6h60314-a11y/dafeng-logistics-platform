# pages/20_撥貨差異.py
import io
import os
import pandas as pd
import streamlit as st

from common_ui import inject_logistics_theme, set_page, card_open, card_close


st.set_page_config(page_title="大豐物流 - 撥貨差異", page_icon="🔁", layout="wide")
inject_logistics_theme()


# =========================
# 讀檔：從 bytes 讀 (可部署)
# =========================
def _read_source_with_lookup_from_bytes(data: bytes):
    """來源檔：回傳 main_df + (若有 '儲位' sheet 則回傳 lookup_df)"""
    head = data[:4096]

    # xlsx
    if head.startswith(b"PK\x03\x04"):
        bio = io.BytesIO(data)
        xls = pd.ExcelFile(bio, engine="openpyxl")
        sheets = xls.sheet_names
        main_sheet = "儲位明細" if "儲位明細" in sheets else sheets[0]
        main_df = pd.read_excel(xls, sheet_name=main_sheet, dtype=str).dropna(how="all").copy()

        lookup_df = None
        if "儲位" in sheets:
            lookup_df = pd.read_excel(xls, sheet_name="儲位", dtype=str).dropna(how="all").copy()

        return main_df, lookup_df

    # xls (OLE2) —— 需要 xlrd
    if head.startswith(b"\xD0\xCF\x11\xe0\xa1\xb1\x1a\xe1"):
        bio = io.BytesIO(data)
        xls = pd.ExcelFile(bio, engine="xlrd")
        sheets = xls.sheet_names
        main_sheet = "儲位明細" if "儲位明細" in sheets else sheets[0]
        main_df = pd.read_excel(xls, sheet_name=main_sheet, dtype=str).dropna(how="all").copy()

        lookup_df = None
        if "儲位" in sheets:
            lookup_df = pd.read_excel(xls, sheet_name="儲位", dtype=str).dropna(how="all").copy()

        return main_df, lookup_df

    # HTML
    head_text = head.decode("utf-8", errors="ignore").lower()
    if "<html" in head_text or "<table" in head_text:
        tables = pd.read_html(io.BytesIO(data), encoding="utf-8", keep_default_na=False)
        if not tables:
            raise ValueError("HTML 內沒有表格可讀取")
        return tables[0].dropna(how="all").copy(), None

    # TSV/CSV（假 xls 常見）
    sample = head.decode("utf-8", errors="ignore")
    sep = "\t" if sample.count("\t") >= sample.count(",") else ","

    last_err = None
    for enc in ("utf-8-sig", "cp950", "big5", "utf-8", "latin1"):
        try:
            txt = data.decode(enc, errors="strict")
            df = pd.read_csv(
                io.StringIO(txt), sep=sep, engine="python",
                dtype=str, keep_default_na=False
            ).dropna(how="all").copy()

            if df.shape[1] <= 1:
                alt_sep = "," if sep == "\t" else "\t"
                df2 = pd.read_csv(
                    io.StringIO(txt), sep=alt_sep, engine="python",
                    dtype=str, keep_default_na=False
                ).dropna(how="all").copy()
                if df2.shape[1] > df.shape[1]:
                    df = df2

            return df, None
        except Exception as e:
            last_err = e
            continue

    raise ValueError(f"無法讀取來源檔。最後錯誤：{last_err}")


def _read_any_table_from_bytes(data: bytes) -> pd.DataFrame:
    """主檔：讀第一張表（支援 xlsx/xls/html/text）"""
    head = data[:4096]

    if head.startswith(b"PK\x03\x04"):
        bio = io.BytesIO(data)
        xls = pd.ExcelFile(bio, engine="openpyxl")
        return pd.read_excel(xls, sheet_name=xls.sheet_names[0], dtype=str).dropna(how="all").copy()

    if head.startswith(b"\xD0\xCF\x11\xe0\xa1\xb1\x1a\xe1"):
        bio = io.BytesIO(data)
        xls = pd.ExcelFile(bio, engine="xlrd")
        return pd.read_excel(xls, sheet_name=xls.sheet_names[0], dtype=str).dropna(how="all").copy()

    head_text = head.decode("utf-8", errors="ignore").lower()
    if "<html" in head_text or "<table" in head_text:
        tables = pd.read_html(io.BytesIO(data), encoding="utf-8", keep_default_na=False)
        if not tables:
            raise ValueError("HTML 內沒有表格可讀取")
        return tables[0].dropna(how="all").copy()

    sample = head.decode("utf-8", errors="ignore")
    sep = "\t" if sample.count("\t") >= sample.count(",") else ","

    last_err = None
    for enc in ("utf-8-sig", "cp950", "big5", "utf-8", "latin1"):
        try:
            txt = data.decode(enc, errors="strict")
            df = pd.read_csv(
                io.StringIO(txt), sep=sep, engine="python",
                dtype=str, keep_default_na=False
            ).dropna(how="all").copy()

            if df.shape[1] <= 1:
                alt_sep = "," if sep == "\t" else "\t"
                df2 = pd.read_csv(
                    io.StringIO(txt), sep=alt_sep, engine="python",
                    dtype=str, keep_default_na=False
                ).dropna(how="all").copy()
                if df2.shape[1] > df.shape[1]:
                    df = df2

            return df
        except Exception as e:
            last_err = e
            continue

    raise ValueError(f"無法讀取主檔。最後錯誤：{last_err}")


# =========================
# 工具（沿用你原本邏輯）
# =========================
def to_num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce")


def is_zero_like(series: pd.Series) -> pd.Series:
    s_str = series.astype(str).str.strip()
    s_num = to_num(series)
    return (s_str == "0") | (s_num == 0)


def find_col_ci(df: pd.DataFrame, target: str):
    t = target.strip().upper()
    for c in df.columns:
        if str(c).strip().upper() == t:
            return c
    return None


def must_col_ci(df: pd.DataFrame, target: str):
    c = find_col_ci(df, target)
    if c is None:
        raise ValueError(f"找不到欄位：{target}")
    return c


def norm_loc(x) -> str:
    if pd.isna(x):
        return ""
    return str(x).strip()[:9]


def reorder_final_cols(df: pd.DataFrame) -> pd.DataFrame:
    cols = ["SONO", "差異量", "儲位", "國際條碼", "商品名稱", "棚別", "商品號", "來源檔名"]
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    return df.loc[:, cols].copy()


def step2_filter_or(df: pd.DataFrame) -> pd.DataFrame:
    c_dif = must_col_ci(df, "AllDIF")
    c_act = must_col_ci(df, "ALLACT")
    cond = (to_num(df[c_dif]) >= 15) | (is_zero_like(df[c_act]))
    return df.loc[cond].copy()


def step3_apply_macro2_logic_keep_sono(df_detail: pd.DataFrame, lookup_df: pd.DataFrame | None) -> pd.DataFrame:
    if df_detail.shape[1] < 16:
        raise ValueError(f"明細欄位不足：目前 {df_detail.shape[1]} 欄，至少要 A~P（16欄）")

    c_sono = find_col_ci(df_detail, "SONO")
    sono_series = df_detail[c_sono] if c_sono is not None else pd.Series([""] * len(df_detail), index=df_detail.index)

    base = df_detail.iloc[:, :16].copy()

    colE = base.columns[4]     # 商品號
    colF = base.columns[5]     # 國際條碼
    colG = base.columns[6]     # 商品名稱
    colH = base.columns[7]     # 差異量
    colP = base.columns[15]    # 儲位

    base[colF] = base[colF].astype(str).str.strip().str[:13]
    base[colP] = base[colP].astype(str).str.strip().str[:9]

    mask_keep = ~is_zero_like(base[colH])
    base = base.loc[mask_keep].copy()
    sono_series = sono_series.loc[base.index]

    shelf = pd.Series([""] * len(base), index=base.index, dtype=object)
    if lookup_df is not None and lookup_df.shape[1] >= 2:
        k = lookup_df.columns[0]
        v = lookup_df.columns[1]
        mapping = dict(
            zip(
                lookup_df[k].astype(str).str.strip().map(norm_loc),
                lookup_df[v].astype(str).str.strip()
            )
        )
        shelf = base[colP].map(norm_loc).map(mapping).fillna("")

    out = pd.DataFrame({
        "SONO": sono_series.astype(str),
        "儲位": base[colP].astype(str).str.strip(),
        "棚別": shelf.astype(str),
        "商品號": base[colE],
        "國際條碼": base[colF],
        "商品名稱": base[colG],
        "差異量": base[colH],
    }).reset_index(drop=True)

    return out


def build_master_loc_shelf_map(df_master: pd.DataFrame) -> dict:
    c_loc = find_col_ci(df_master, "儲位")
    c_shelf = find_col_ci(df_master, "棚別")

    if c_loc is None or c_shelf is None:
        if df_master.shape[1] < 2:
            raise ValueError("主檔欄位不足，至少要兩欄（儲位、棚別）")
        c_loc = df_master.columns[0]
        c_shelf = df_master.columns[1]

    loc = df_master[c_loc].map(norm_loc)
    shelf = df_master[c_shelf].astype(str).str.strip()
    m = (loc != "")
    return dict(zip(loc[m], shelf[m]))


def step4_overwrite_shelf(df_macro2: pd.DataFrame, master_map: dict) -> tuple[pd.DataFrame, pd.DataFrame]:
    df = df_macro2.copy()
    key = df["儲位"].map(norm_loc)
    shelf_master = key.map(master_map)

    df["棚別"] = shelf_master.combine_first(df["棚別"]).fillna("")
    not_found = df[shelf_master.isna()].copy()
    return df, not_found


def main():
    set_page("撥貨差異（棚別比對）", icon="🔁", subtitle="多來源檔 OR 篩選 → 第二巨集邏輯 → 主檔棚別覆蓋 → 匯出下載")

    card_open("📦 出貨課｜撥貨差異分析")
    st.write("① 上傳來源檔（可多選，含 AllDIF / ALLACT / 儲位明細）")
    src_files = st.file_uploader(
        "來源檔（多選）",
        type=["xlsx", "xls", "csv", "tsv", "txt", "htm", "html"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )

    st.write("② 上傳主檔（儲位→棚別）")
    master_file = st.file_uploader(
        "主檔（單檔）",
        type=["xlsx", "xls", "csv", "tsv", "txt", "htm", "html"],
        accept_multiple_files=False,
        label_visibility="collapsed",
    )

    run = st.button("🚀 開始分析並產出 Excel", use_container_width=True, disabled=not (src_files and master_file))

    if run:
        try:
            with st.spinner("讀取主檔中..."):
                df_master = _read_any_table_from_bytes(master_file.getvalue())
                master_map = build_master_loc_shelf_map(df_master)

            all_or, all_final, all_notfound = [], [], []
            prog = st.progress(0)
            total = len(src_files)

            for i, uf in enumerate(src_files, 1):
                base_name = uf.name

                df_src, lookup_df = _read_source_with_lookup_from_bytes(uf.getvalue())
                df_detail = step2_filter_or(df_src)
                df_detail = df_detail.copy()
                df_detail["來源檔名"] = base_name
                all_or.append(df_detail)

                df_macro2 = step3_apply_macro2_logic_keep_sono(df_detail.drop(columns=["來源檔名"]), lookup_df)
                df_macro2["來源檔名"] = base_name

                df_final, df_notfound = step4_overwrite_shelf(df_macro2, master_map)
                df_final["來源檔名"] = base_name
                df_notfound["來源檔名"] = base_name

                df_final = reorder_final_cols(df_final)
                all_final.append(df_final)

                if not df_notfound.empty:
                    df_notfound = reorder_final_cols(df_notfound)
                    all_notfound.append(df_notfound)

                prog.progress(int(i / total * 100))

            out_or = pd.concat(all_or, ignore_index=True) if all_or else pd.DataFrame()
            out_final = pd.concat(all_final, ignore_index=True) if all_final else pd.DataFrame()
            out_notfound = pd.concat(all_notfound, ignore_index=True) if all_notfound else pd.DataFrame()

            out_name = os.path.splitext(src_files[0].name)[0] + "_多檔_棚別覆蓋.xlsx"

            bio = io.BytesIO()
            with pd.ExcelWriter(bio, engine="openpyxl") as writer:
                (out_or if not out_or.empty else pd.DataFrame({"msg": ["無資料"]})).to_excel(
                    writer, index=False, sheet_name="1_篩選明細_OR"
                )
                (out_final if not out_final.empty else pd.DataFrame({"msg": ["無資料"]})).to_excel(
                    writer, index=False, sheet_name="差異明細"
                )
                (out_notfound if not out_notfound.empty else pd.DataFrame({"msg": ["無資料"]})).to_excel(
                    writer, index=False, sheet_name="3_主檔找不到儲位"
                )

            bio.seek(0)
            st.success("✅ 已完成！請下載結果檔。")
            st.download_button(
                "⬇️ 下載輸出 Excel",
                data=bio,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        except Exception as e:
            st.error(f"❌ 執行失敗：{e}")

    card_close()


if __name__ == "__main__":
    main()
