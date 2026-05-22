import io
from datetime import datetime

import pandas as pd
import streamlit as st

from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont

try:
    from common_ui import set_page, inject_logistics_theme, card_open, card_close
except Exception:
    set_page = None
    inject_logistics_theme = None
    card_open = None
    card_close = None


# =========================
# 頁面設定
# =========================
if set_page:
    set_page("客訂差異分析", "出貨課")
else:
    st.set_page_config(page_title="客訂差異分析", layout="wide")

if inject_logistics_theme:
    inject_logistics_theme()

st.title("30｜客訂差異分析")
st.caption("原始明細 × 商品主檔 × 其他儲位 × 儲位明細，自動產出最後篩選明細。")


# =========================
# 欄位名稱設定
# =========================
TARGET_HEADER = "原始配庫存量"
DIFF_HEADER_1 = "差異量"
DIFF_HEADER_2 = "原始配庫差異量"
QTY_HEADER_1 = "計量單位數量"
QTY_HEADER_2 = "數量"
EXCLUDE_HEADER = "成箱箱號"
ORDER_HEADER = "貨主訂單"
REMARK_HEADER = "備註"
CUSTOM_ORDER_HEADER = "客訂"
LOCATION_HEADER = "儲位"
ITEM_HEADER = "商品"
BATCH_HEADER = "揀貨批次號"
MASTER_ITEM_HEADER = "商品代號"
BARCODE_HEADER = "國際條碼"
PRODUCT_NAME_HEADER = "名稱"
OTHER_ITEM_HEADER = "商品號"
OTHER_LOCATION_HEADER = "儲位"
OTHER_EXPIRY_HEADER = "商品效期"
OTHER_DAYS_HEADER = "效期剩餘天數"
LOCATION_DETAIL_LOCATION_HEADER = "儲位"
SHED_HEADER = "棚別"
PIVOT_VALUE_HEADER = "加總 - 原始配庫差異量"
PIVOT_CUSTOM_HEADER = "是否為客訂"
FINAL_REMARK_HEADER = "備註"
DETAIL_SHEET_NAME = "完整明細"
NEW_SHEET_NAME = "指定欄位"
FINAL_SHEET_NAME = "最後篩選明細"


# =========================
# 共用工具
# =========================
def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [str(col).strip() for col in df.columns]
    return df


def has_value_series(series: pd.Series) -> pd.Series:
    return series.notna() & (series.astype(str).str.strip() != "")


def to_number_series(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.astype(str).str.replace(",", "", regex=False).str.strip(),
        errors="coerce",
    ).fillna(0)


def clean_text_series(series: pd.Series, upper: bool = False) -> pd.Series:
    s = series.fillna("").astype(str).str.strip()
    if upper:
        s = s.str.upper()
    return s


def is_custom_order_series(series: pd.Series) -> pd.Series:
    """判斷貨主訂單是否為客訂：開頭為 GB、U、20XX。"""
    order_text = clean_text_series(series, upper=True)
    return (
        order_text.str.startswith("GB")
        | order_text.str.startswith("U")
        | order_text.str.match(r"^20\d{2}")
    )


def read_excel_upload(uploaded_file, label: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(uploaded_file, dtype=str)
        return clean_columns(df)
    except ImportError as e:
        raise RuntimeError(f"{label} 讀取失敗：若上傳 .xls，請在 requirements.txt 加上 xlrd。原始錯誤：{e}")
    except Exception as e:
        raise RuntimeError(f"{label} 讀取失敗：{e}")


def check_required_columns(df: pd.DataFrame, required_columns: list[str], label: str) -> None:
    missing = [col for col in required_columns if col not in df.columns]
    if missing:
        raise ValueError(f"{label} 找不到以下欄位：" + "、".join(missing))


# =========================
# 第 2 個檔案：商品主檔
# =========================
def read_product_master(master_file) -> pd.DataFrame:
    df_master = read_excel_upload(master_file, "商品主檔")
    required_columns = [MASTER_ITEM_HEADER, BARCODE_HEADER, PRODUCT_NAME_HEADER]
    check_required_columns(df_master, required_columns, "商品主檔")

    df_master = df_master[[MASTER_ITEM_HEADER, BARCODE_HEADER, PRODUCT_NAME_HEADER]].copy()
    df_master[MASTER_ITEM_HEADER] = clean_text_series(df_master[MASTER_ITEM_HEADER])
    df_master[BARCODE_HEADER] = clean_text_series(df_master[BARCODE_HEADER])
    df_master[PRODUCT_NAME_HEADER] = clean_text_series(df_master[PRODUCT_NAME_HEADER])
    df_master = df_master[df_master[MASTER_ITEM_HEADER] != ""].copy()
    df_master = df_master.drop_duplicates(subset=[MASTER_ITEM_HEADER], keep="first")
    return df_master.rename(columns={MASTER_ITEM_HEADER: ITEM_HEADER})


def add_product_info(df: pd.DataFrame, df_master: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for col in [BARCODE_HEADER, PRODUCT_NAME_HEADER]:
        if col in df.columns:
            df = df.drop(columns=[col])
    df[ITEM_HEADER] = clean_text_series(df[ITEM_HEADER])
    df = df.merge(df_master, on=ITEM_HEADER, how="left")
    df[BARCODE_HEADER] = df[BARCODE_HEADER].fillna("")
    df[PRODUCT_NAME_HEADER] = df[PRODUCT_NAME_HEADER].fillna("")
    return df


# =========================
# 第 3 個檔案：其他儲位 / 庫存明細
# =========================
def read_other_location_file(other_file) -> pd.DataFrame:
    df_other = read_excel_upload(other_file, "其他儲位 / 庫存明細")
    required_columns = [OTHER_ITEM_HEADER, OTHER_LOCATION_HEADER, OTHER_EXPIRY_HEADER, OTHER_DAYS_HEADER]
    check_required_columns(df_other, required_columns, "其他儲位 / 庫存明細")

    df_other = df_other[[OTHER_ITEM_HEADER, OTHER_LOCATION_HEADER, OTHER_EXPIRY_HEADER, OTHER_DAYS_HEADER]].copy()
    df_other[OTHER_ITEM_HEADER] = clean_text_series(df_other[OTHER_ITEM_HEADER])
    df_other[OTHER_LOCATION_HEADER] = clean_text_series(df_other[OTHER_LOCATION_HEADER])
    df_other[OTHER_EXPIRY_HEADER] = clean_text_series(df_other[OTHER_EXPIRY_HEADER])
    df_other[OTHER_DAYS_HEADER] = to_number_series(df_other[OTHER_DAYS_HEADER])
    df_other = df_other[(df_other[OTHER_ITEM_HEADER] != "") & (df_other[OTHER_LOCATION_HEADER] != "")].copy()

    # 效期越短越優先
    return df_other.sort_values(
        by=[OTHER_ITEM_HEADER, OTHER_DAYS_HEADER, OTHER_EXPIRY_HEADER, OTHER_LOCATION_HEADER],
        ascending=[True, True, True, True],
    )


def add_other_locations_to_final(df_final: pd.DataFrame, df_other: pd.DataFrame) -> pd.DataFrame:
    df_final = df_final.copy()
    other_lookup = {}

    for item_code, group in df_other.groupby(OTHER_ITEM_HEADER):
        locations = []
        for _, row in group.iterrows():
            loc = str(row[OTHER_LOCATION_HEADER]).strip()
            if loc and loc not in locations:
                locations.append(loc)
        other_lookup[str(item_code).strip()] = locations

    loc1_list, loc2_list, loc3_list = [], [], []
    for _, row in df_final.iterrows():
        item = str(row.get("商品", "")).strip()
        current_loc = str(row.get("儲位", "")).strip()
        other_locations = [loc for loc in other_lookup.get(item, []) if loc != current_loc]
        loc1_list.append(other_locations[0] if len(other_locations) >= 1 else "")
        loc2_list.append(other_locations[1] if len(other_locations) >= 2 else "")
        loc3_list.append(other_locations[2] if len(other_locations) >= 3 else "")

    df_final["儲位1"] = loc1_list
    df_final["儲位2"] = loc2_list
    df_final["儲位3"] = loc3_list
    return df_final


# =========================
# 第 4 個檔案：儲位明細
# =========================
def read_location_detail_file(location_detail_file) -> pd.DataFrame:
    df_location = read_excel_upload(location_detail_file, "儲位明細")
    required_columns = [LOCATION_DETAIL_LOCATION_HEADER, SHED_HEADER]
    check_required_columns(df_location, required_columns, "儲位明細")

    df_location = df_location[[LOCATION_DETAIL_LOCATION_HEADER, SHED_HEADER]].copy()
    df_location[LOCATION_DETAIL_LOCATION_HEADER] = clean_text_series(df_location[LOCATION_DETAIL_LOCATION_HEADER])
    df_location[SHED_HEADER] = clean_text_series(df_location[SHED_HEADER])
    df_location = df_location[df_location[LOCATION_DETAIL_LOCATION_HEADER] != ""].copy()
    return df_location.drop_duplicates(subset=[LOCATION_DETAIL_LOCATION_HEADER], keep="first")


def add_shed_to_final(df_final: pd.DataFrame, df_location: pd.DataFrame) -> pd.DataFrame:
    df_final = df_final.copy()
    if SHED_HEADER in df_final.columns:
        df_final = df_final.drop(columns=[SHED_HEADER])

    df_final["儲位"] = clean_text_series(df_final["儲位"])
    df_final = df_final.merge(
        df_location,
        left_on="儲位",
        right_on=LOCATION_DETAIL_LOCATION_HEADER,
        how="left",
    )

    if LOCATION_DETAIL_LOCATION_HEADER in df_final.columns and LOCATION_DETAIL_LOCATION_HEADER != "儲位":
        df_final = df_final.drop(columns=[LOCATION_DETAIL_LOCATION_HEADER])

    df_final[SHED_HEADER] = df_final[SHED_HEADER].fillna("")
    return df_final


# =========================
# 主處理流程
# =========================
def process_excel(main_file, master_file, other_file, location_detail_file, batch_input: str, min_diff_value):
    logs = []

    def log(msg: str):
        logs.append(msg)

    df = read_excel_upload(main_file, "原始明細")
    original_count = len(df)

    required_columns = [
        TARGET_HEADER,
        QTY_HEADER_1,
        QTY_HEADER_2,
        EXCLUDE_HEADER,
        ORDER_HEADER,
        LOCATION_HEADER,
        ITEM_HEADER,
        BATCH_HEADER,
    ]
    check_required_columns(df, required_columns, "原始明細")

    df_master = read_product_master(master_file)
    df_other = read_other_location_file(other_file)
    df_location = read_location_detail_file(location_detail_file)

    log(f"原始筆數：{original_count}")
    log(f"商品主檔筆數：{len(df_master)}")
    log(f"其他儲位檔筆數：{len(df_other)}")
    log(f"儲位明細檔筆數：{len(df_location)}")

    # 1. 排除「成箱箱號」有值的明細
    df = df[~has_value_series(df[EXCLUDE_HEADER])].copy()
    after_box_filter_count = len(df)
    deleted_box_count = original_count - after_box_filter_count
    log(f"已排除成箱箱號有值筆數：{deleted_box_count}")
    log(f"成箱箱號排除後保留筆數：{after_box_filter_count}")

    # 2. 差異量 = 計量單位數量 - 數量
    diff_1 = to_number_series(df[QTY_HEADER_1]) - to_number_series(df[QTY_HEADER_2])

    # 3. 原始配庫差異量 = 原始配庫存量 - 數量
    diff_2 = to_number_series(df[TARGET_HEADER]) - to_number_series(df[QTY_HEADER_2])

    for col in [DIFF_HEADER_1, DIFF_HEADER_2]:
        if col in df.columns:
            df = df.drop(columns=[col])

    target_index = df.columns.get_loc(TARGET_HEADER)
    df.insert(target_index + 1, DIFF_HEADER_1, diff_1)
    df.insert(target_index + 2, DIFF_HEADER_2, diff_2)

    # 4. 刪除「原始配庫差異量」= 0 的明細
    before_diff_filter_count = len(df)
    df = df[to_number_series(df[DIFF_HEADER_2]) != 0].copy()
    after_diff_filter_count = len(df)
    deleted_zero_diff_count = before_diff_filter_count - after_diff_filter_count
    log(f"已刪除原始配庫差異量=0筆數：{deleted_zero_diff_count}")
    log(f"刪除0後保留筆數：{after_diff_filter_count}")

    # 5. 在「貨主訂單」右邊新增「備註」與「客訂」
    custom_order_value = is_custom_order_series(df[ORDER_HEADER]).map(lambda x: "客訂" if x else "")
    for col in [REMARK_HEADER, CUSTOM_ORDER_HEADER]:
        if col in df.columns:
            df = df.drop(columns=[col])

    order_index = df.columns.get_loc(ORDER_HEADER)
    df.insert(order_index + 1, REMARK_HEADER, custom_order_value)
    df.insert(order_index + 2, CUSTOM_ORDER_HEADER, custom_order_value)
    custom_order_count = (df[CUSTOM_ORDER_HEADER] == "客訂").sum()
    log(f"客訂筆數：{custom_order_count}")

    # 6. 帶出商品主檔資訊
    df = add_product_info(df, df_master)
    matched_barcode_count = (df[BARCODE_HEADER].astype(str).str.strip() != "").sum()
    matched_name_count = (df[PRODUCT_NAME_HEADER].astype(str).str.strip() != "").sum()
    log(f"對應到國際條碼筆數：{matched_barcode_count}")
    log(f"對應到名稱筆數：{matched_name_count}")

    # 7. 複製指定欄位到新的工作表
    copy_columns = [
        LOCATION_HEADER,
        ITEM_HEADER,
        BARCODE_HEADER,
        PRODUCT_NAME_HEADER,
        DIFF_HEADER_2,
        CUSTOM_ORDER_HEADER,
        BATCH_HEADER,
    ]
    check_required_columns(df, copy_columns, "要複製到指定欄位的資料")
    df_new_sheet = df[copy_columns].copy()
    df_new_sheet[DIFF_HEADER_2] = to_number_series(df_new_sheet[DIFF_HEADER_2])
    df_new_sheet[BATCH_HEADER] = clean_text_series(df_new_sheet[BATCH_HEADER], upper=True)
    df_new_sheet[ITEM_HEADER] = clean_text_series(df_new_sheet[ITEM_HEADER])
    df_new_sheet[LOCATION_HEADER] = clean_text_series(df_new_sheet[LOCATION_HEADER])
    df_new_sheet[CUSTOM_ORDER_HEADER] = clean_text_series(df_new_sheet[CUSTOM_ORDER_HEADER])
    df_new_sheet[BARCODE_HEADER] = clean_text_series(df_new_sheet[BARCODE_HEADER])
    df_new_sheet[PRODUCT_NAME_HEADER] = clean_text_series(df_new_sheet[PRODUCT_NAME_HEADER])
    log(f"新工作表「{NEW_SHEET_NAME}」明細筆數：{len(df_new_sheet)}")

    # 8. 揀貨批次號
    batch_list = [x.strip().upper() for x in str(batch_input).replace("，", ",").split(",") if x.strip()]
    if not batch_list:
        raise ValueError("請輸入至少一個揀貨批次號。")

    log(f"輸入揀貨批次號：{', '.join(batch_list)}")
    df_pivot_source = df_new_sheet[df_new_sheet[BATCH_HEADER].isin(batch_list)].copy()
    pivot_source_count = len(df_pivot_source)
    if pivot_source_count == 0:
        raise ValueError("找不到你輸入的揀貨批次號資料：" + "、".join(batch_list))
    log(f"指定揀貨批次號篩選後筆數：{pivot_source_count}")

    # 9. 樞紐分析
    df_pivot = (
        df_pivot_source.groupby(
            [LOCATION_HEADER, ITEM_HEADER, BARCODE_HEADER, PRODUCT_NAME_HEADER],
            dropna=False,
            as_index=False,
        )[DIFF_HEADER_2]
        .sum()
    )
    df_pivot = df_pivot.rename(
        columns={
            LOCATION_HEADER: "列標籤",
            ITEM_HEADER: "商品",
            DIFF_HEADER_2: PIVOT_VALUE_HEADER,
        }
    )
    df_pivot = df_pivot.sort_values(by=["列標籤", "商品"], ascending=[True, True])

    # 10. 是否為客訂
    custom_lookup = df_new_sheet[[ITEM_HEADER, CUSTOM_ORDER_HEADER]].copy()
    custom_lookup[ITEM_HEADER] = clean_text_series(custom_lookup[ITEM_HEADER])
    custom_lookup[CUSTOM_ORDER_HEADER] = clean_text_series(custom_lookup[CUSTOM_ORDER_HEADER])
    custom_lookup = custom_lookup.groupby(ITEM_HEADER, as_index=False)[CUSTOM_ORDER_HEADER].agg(
        lambda x: "客訂" if (x == "客訂").any() else "0"
    )
    custom_lookup_dict = dict(zip(custom_lookup[ITEM_HEADER], custom_lookup[CUSTOM_ORDER_HEADER]))
    df_pivot["商品"] = clean_text_series(df_pivot["商品"])
    df_pivot[PIVOT_CUSTOM_HEADER] = df_pivot["商品"].map(custom_lookup_dict).fillna("0")
    log(f"樞紐分析筆數：{len(df_pivot)}")

    # 11. 篩選條件
    if min_diff_value is None:
        df_final = df_pivot[df_pivot[PIVOT_CUSTOM_HEADER] == "客訂"].copy()
        final_filter_desc = "備註 = 客訂"
        min_diff_text = "未輸入"
    else:
        df_final = df_pivot[
            (df_pivot[PIVOT_CUSTOM_HEADER] == "客訂")
            | (to_number_series(df_pivot[PIVOT_VALUE_HEADER]) >= float(min_diff_value))
        ].copy()
        final_filter_desc = f"備註 = 客訂，或 {PIVOT_VALUE_HEADER} >= {min_diff_value}"
        min_diff_text = str(min_diff_value).replace(".", "_")

    # 12. 調整「最後篩選明細」欄位名稱
    df_final = df_final.rename(
        columns={
            PIVOT_VALUE_HEADER: "差異量",
            "列標籤": "儲位",
            PIVOT_CUSTOM_HEADER: FINAL_REMARK_HEADER,
        }
    )

    # 13. 比對其他儲位，依短效優先帶出 儲位1、儲位2、儲位3
    df_final = add_other_locations_to_final(df_final, df_other)

    # 14. 比對儲位明細，帶出棚別
    df_final = add_shed_to_final(df_final, df_location)

    # 15. 最終欄位順序：棚別移到最右欄
    final_columns_order = [
        "差異量",
        "儲位",
        BARCODE_HEADER,
        PRODUCT_NAME_HEADER,
        "商品",
        FINAL_REMARK_HEADER,
        "儲位1",
        "儲位2",
        "儲位3",
        SHED_HEADER,
    ]
    final_columns_order = [col for col in final_columns_order if col in df_final.columns]
    df_final = df_final[final_columns_order].copy()
    final_count = len(df_final)
    log(f"最後篩選明細筆數：{final_count}")
    log(f"最後完整明細筆數：{len(df)}")
    log(f"最後篩選條件：{final_filter_desc}")

    # 16. 輸出 Excel 到記憶體
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=DETAIL_SHEET_NAME, index=False)
        df_new_sheet.to_excel(writer, sheet_name=NEW_SHEET_NAME, index=False, startrow=0, startcol=0)
        df_pivot.to_excel(writer, sheet_name=NEW_SHEET_NAME, index=False, startrow=0, startcol=8)
        df_final.to_excel(writer, sheet_name=FINAL_SHEET_NAME, index=False)

        workbook = writer.book
        ws_detail = workbook[DETAIL_SHEET_NAME]
        ws_pivot = workbook[NEW_SHEET_NAME]
        ws_final = workbook[FINAL_SHEET_NAME]

        for ws in [ws_detail, ws_pivot, ws_final]:
            for col in ws.columns:
                max_length = 0
                col_letter = col[0].column_letter
                for cell in col:
                    value = cell.value
                    if value is not None:
                        max_length = max(max_length, len(str(value)))
                ws.column_dimensions[col_letter].width = min(max_length + 2, 35)
            ws.freeze_panes = "A2"

        # 最後篩選明細：「國際條碼」後五碼放大加粗
        barcode_col_index = None
        for cell in ws_final[1]:
            if str(cell.value).strip() == BARCODE_HEADER:
                barcode_col_index = cell.column
                break

        if barcode_col_index is not None:
            normal_font = InlineFont(sz=11)
            big_font = InlineFont(sz=18, b=True)
            for row in range(2, ws_final.max_row + 1):
                cell = ws_final.cell(row=row, column=barcode_col_index)
                barcode_value = cell.value
                if barcode_value is None:
                    continue
                barcode_text = str(barcode_value).strip()
                if barcode_text == "":
                    continue
                if len(barcode_text) > 5:
                    cell.value = CellRichText(
                        TextBlock(normal_font, barcode_text[:-5]),
                        TextBlock(big_font, barcode_text[-5:]),
                    )
                else:
                    cell.value = CellRichText(TextBlock(big_font, barcode_text))
                ws_final.row_dimensions[row].height = 24

    output.seek(0)

    result = {
        "output_bytes": output.getvalue(),
        "logs": logs,
        "summary": {
            "原始筆數": original_count,
            "排除成箱箱號有值筆數": deleted_box_count,
            "刪除原始配庫差異量=0筆數": deleted_zero_diff_count,
            "最後完整明細筆數": len(df),
            "商品主檔對應到國際條碼筆數": int(matched_barcode_count),
            "商品主檔對應到名稱筆數": int(matched_name_count),
            "批次篩選後筆數": pivot_source_count,
            "樞紐分析筆數": len(df_pivot),
            "最後篩選明細筆數": final_count,
            "最後篩選條件": final_filter_desc,
            "門檻文字": min_diff_text,
        },
        "df_final": df_final,
    }
    return result


# =========================
# Streamlit UI
# =========================
with st.container():
    if card_open:
        card_open("檔案上傳")
    st.subheader("1. 上傳檔案")
    c1, c2 = st.columns(2)
    with c1:
        main_file = st.file_uploader("第 1 個檔案：原始明細 Excel", type=["xlsx", "xls"], key="main_file")
        other_file = st.file_uploader("第 3 個檔案：其他儲位 / 庫存明細 Excel", type=["xlsx", "xls"], key="other_file")
    with c2:
        master_file = st.file_uploader("第 2 個檔案：商品主檔 Excel", type=["xlsx", "xls"], key="master_file")
        location_detail_file = st.file_uploader("第 4 個檔案：儲位明細 Excel", type=["xlsx", "xls"], key="location_detail_file")
    if card_close:
        card_close()

with st.container():
    if card_open:
        card_open("執行條件")
    st.subheader("2. 輸入篩選條件")
    batch_input = st.text_input(
        "揀貨批次號",
        placeholder="可輸入單一批次，例如 C20315A02；多個批次請用逗號分隔",
    )
    use_threshold = st.checkbox("除了客訂之外，也保留差異量達門檻的資料", value=False)
    min_diff_value = None
    if use_threshold:
        min_diff_value = st.number_input(
            "差異量門檻：加總 - 原始配庫差異量 >=",
            min_value=0.0,
            value=2.0,
            step=1.0,
        )
    if card_close:
        card_close()

ready = all([main_file, master_file, other_file, location_detail_file, batch_input.strip()])
run_button = st.button("開始產生客訂差異報表", type="primary", disabled=not ready, use_container_width=True)

if not ready:
    st.info("請先上傳 4 個 Excel，並輸入揀貨批次號。")

if run_button:
    try:
        with st.spinner("資料處理中，請稍候..."):
            result = process_excel(
                main_file=main_file,
                master_file=master_file,
                other_file=other_file,
                location_detail_file=location_detail_file,
                batch_input=batch_input,
                min_diff_value=min_diff_value if use_threshold else None,
            )

        st.success("處理完成，可以下載 Excel。")

        summary = result["summary"]
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("原始筆數", summary["原始筆數"])
        m2.metric("批次篩選後", summary["批次篩選後筆數"])
        m3.metric("樞紐分析筆數", summary["樞紐分析筆數"])
        m4.metric("最後篩選明細", summary["最後篩選明細筆數"])

        st.markdown("#### 處理紀錄")
        st.code("\n".join(result["logs"]), language="text")

        st.markdown("#### 最後篩選明細預覽")
        st.dataframe(result["df_final"], use_container_width=True)

        now_text = datetime.now().strftime("%Y%m%d_%H%M")
        safe_batch = "_".join([x.strip().upper() for x in batch_input.replace("，", ",").split(",") if x.strip()])[:80]
        download_name = f"客訂差異_批次_{safe_batch}_{now_text}.xlsx"
        st.download_button(
            label="下載 Excel 報表",
            data=result["output_bytes"],
            file_name=download_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )

    except Exception as e:
        st.error(str(e))
        st.exception(e)
