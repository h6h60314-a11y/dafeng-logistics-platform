# pages/9_大樹KPI首頁.py
import streamlit as st
from urllib.parse import quote, unquote

from common_ui import inject_logistics_theme, set_page, card_open, card_close

st.set_page_config(page_title="大樹KPI", page_icon="📈", layout="wide")
inject_logistics_theme()

# ✅ 允許從 KPI 首頁導頁的模組清單（安全白名單）
ALLOW_PAGES = {
    "pages/10_進貨驗收量.py",
    "pages/11_庫存訂單應出量分析.py",
    "pages/12_越庫訂單分析.py",
    "pages/13_庫存訂單實出量分析.py",
    "pages/14_每日上架分析.py",
    "pages/15_庫存盤點正確率.py",
    "pages/16_門市到貨異常率.py",
    "pages/18_各類儲區使用率.py",  # ✅ 新增
}


def _route_by_query():
    qp = st.query_params
    raw = qp.get("page", "")

    if isinstance(raw, list):
        raw = raw[0] if raw else ""

    if not raw:
        return

    target = unquote(raw)
    st.query_params.clear()

    if target not in ALLOW_PAGES:
        return

    try:
        st.switch_page(target)
    except Exception:
        return


def _css_and_js():
    st.markdown(
        r"""
<style>
.kpi-list{ margin-top: 6px; }
.kpi-row{
  display: flex;
  align-items: flex-start;
  gap: 10px;
  margin: 12px 0;
}
.kpi-left{
  width: 34px;
  flex: 0 0 34px;
  display: inline-flex;
  align-items: flex-start;
  gap: 8px;
  margin-top: 2px;
}
.kpi-bullet{
  color: rgba(15, 23, 42, 0.55);
  font-size: 16px;
  line-height: 1;
}
.kpi-ico{ font-size: 16px; line-height: 1; }
.kpi-right{ flex: 1 1 auto; line-height: 1.55; }
.kpi-link{
  display: inline;
  color: rgba(15, 23, 42, 0.92) !important;
  font-weight: 900;
  font-size: 16px;
  line-height: 1.45;
  text-decoration: none !important;
  cursor: pointer;
}
.kpi-link:hover{ opacity: 0.86; }
.kpi-desc{
  display: inline;
  margin-left: 6px;
  color: rgba(15, 23, 42, 0.72);
  font-weight: 650;
  font-size: 14px;
  line-height: 1.45;
}
div[data-testid="stMarkdown"]{ margin: 0 !important; }
</style>

<script>
(function () {
  function bind() {
    document.querySelectorAll('a.kpi-link').forEach(a => {
      a.addEventListener('click', (e) => {
        e.preventDefault();
        window.location.assign(a.getAttribute('href'));
      }, { passive: false });
    });
  }
  const root = document.querySelector('#root') || document.body;
  const obs = new MutationObserver(() => bind());
  obs.observe(root, { childList: true, subtree: true });
  bind();
})();
</script>
""",
        unsafe_allow_html=True,
    )


def _nav_item(icon: str, title: str, page_path: str, desc: str):
    encoded = quote(page_path, safe="/_.-")
    st.markdown(
        (
            f'<div class="kpi-row">'
            f'  <div class="kpi-left">'
            f'    <span class="kpi-bullet">•</span>'
            f'    <span class="kpi-ico">{icon}</span>'
            f'  </div>'
            f'  <div class="kpi-right">'
            f'    <a class="kpi-link" href="?page={encoded}" target="_self">{title}：</a>'
            f'    <span class="kpi-desc">{desc}</span>'
            f'  </div>'
            f'</div>'
        ),
        unsafe_allow_html=True,
    )


def main():
    _route_by_query()

    set_page("大樹KPI", icon="📈", subtitle="KPI 模組入口｜匯總｜告警｜趨勢")

    card_open("📈 大樹KPI模組")
    _css_and_js()

    st.markdown('<div class="kpi-list">', unsafe_allow_html=True)

    _nav_item(
        "📥",
        "進貨驗收量",
        "pages/10_進貨驗收量.py",
        "GPO / GXPO：供應商、採購單、品號、驗收數量合計",
    )

    _nav_item(
        "📦",
        "庫存訂單應出量分析",
        "pages/11_庫存訂單應出量分析.py",
        "庫存出貨訂單量、總揀（儲位數/品項數），明細預覽與匯出",
    )

    _nav_item(
        "🧾",
        "越庫訂單分析",
        "pages/12_越庫訂單分析.py",
        "兩表比對 CLOSE_USER｜排除 FT03~FT11｜統計越庫應作/實作｜輸出結果",
    )

    _nav_item(
        "🚚",
        "庫存訂單實出量分析",
        "pages/13_庫存訂單實出量分析.py",
        "實際出貨量（零散/成箱/訂單筆數）｜混庫出貨件數（零散/成箱）",
    )

    _nav_item(
        "📦",
        "每日上架分析",
        "pages/14_每日上架分析.py",
        "前一日上架清單：排除指定儲位代碼｜上架筆數與上架總量統計",
    )

    _nav_item(
        "🎯",
        "庫存盤點正確率",
        "pages/15_庫存盤點正確率.py",
        "統計：商品號去重、儲位筆數、差異≠0筆數、正確率、差異正負總和",
    )

    _nav_item(
        "🏪",
        "門市到貨異常率",
        "pages/16_門市到貨異常率.py",
        "依箱號解析年月日｜排除供應商｜統計多貨/短少/凹破漏｜匯出處理後明細",
    )

    # ✅ 新增：18_各類儲區使用率
    _nav_item(
        "🧊",
        "各類儲區使用率",
        "pages/18_各類儲區使用率.py",
        "區(溫層)使用率（有效/已用/未用/使用率）＋棚別分類統計（含未知明細）",
    )

    st.markdown("</div>", unsafe_allow_html=True)
    card_close()


if __name__ == "__main__":
    main()
