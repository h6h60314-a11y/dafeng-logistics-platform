# pages/7_出貨課首頁.py
import streamlit as st

from common_ui import inject_logistics_theme, set_page, card_open, card_close

st.set_page_config(page_title="出貨課", page_icon="📦", layout="wide")
inject_logistics_theme()


def _css():
    st.markdown(
        """
<style>
/* 出貨課首頁模組清單 */
.outbound-list {
    margin-top: 8px;
}

.outbound-item {
    padding: 10px 4px 12px 4px;
    border-bottom: 1px solid rgba(148, 163, 184, 0.22);
}

.outbound-desc {
    margin-top: -4px;
    margin-left: 34px;
    color: rgba(15, 23, 42, 0.68);
    font-size: 14px;
    font-weight: 650;
    line-height: 1.55;
}

/* Streamlit page_link 微調 */
div[data-testid="stPageLink"] {
    margin: 0 !important;
}

div[data-testid="stPageLink"] a {
    font-weight: 900 !important;
    font-size: 16px !important;
    color: rgba(15, 23, 42, 0.94) !important;
    text-decoration: none !important;
}

div[data-testid="stPageLink"] a:hover {
    opacity: 0.86;
    text-decoration: none !important;
}

/* 減少 markdown 間距 */
div[data-testid="stMarkdown"] {
    margin-bottom: 0 !important;
}
</style>
""",
        unsafe_allow_html=True,
    )


def _nav_item(icon: str, title: str, page_path: str, desc: str):
    st.markdown('<div class="outbound-item">', unsafe_allow_html=True)

    st.page_link(
        page_path,
        label=title,
        icon=icon,
    )

    st.markdown(
        f'<div class="outbound-desc">{desc}</div>',
        unsafe_allow_html=True,
    )

    st.markdown("</div>", unsafe_allow_html=True)


def main():
    set_page("出貨課", icon="📦", subtitle="Outbound｜出貨相關模組入口")

    card_open("📦 出貨課模組")
    _css()

    st.markdown('<div class="outbound-list">', unsafe_allow_html=True)

    _nav_item(
        "📦",
        "撥貨差異",
        "pages/6_撥貨差異.py",
        "AllDIF / ALLACT 篩選、明細匯整、儲位比對棚別、輸出差異明細",
    )

    _nav_item(
        "📄",
        "採品門市差異量",
        "pages/23_採品門市差異量.py",
        "依『未配出原因』回填至同名分頁，輸出更新後差異量檔",
    )

    _nav_item(
        "📦",
        "出貨作業線產能",
        "pages/24_出貨作業線產能.py",
        "每日各作業線產力狀況，輸出每日達標檔",
    )

    _nav_item(
        "⏱️",
        "各時段作業效率",
        "pages/29_各時段作業效率.py",
        "依 LINEID / ZONEID 1~4 計算去重後加權 PCS，逐時段 PASS / FAIL 著色並可下載 Excel",
    )

    _nav_item(
        "🧾",
        "客訂差異",
        "pages/30_客訂差異.py",
        "客訂差異篩選、商品主檔比對、其他儲位短效優先、棚別回填並輸出最後篩選明細",
    )

    st.markdown("</div>", unsafe_allow_html=True)

    card_close()


if __name__ == "__main__":
    main()
