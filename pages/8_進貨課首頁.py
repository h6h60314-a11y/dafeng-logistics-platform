# pages/8_進貨課首頁.py
import streamlit as st
from urllib.parse import quote, unquote

from common_ui import inject_logistics_theme, set_page, card_open, card_close

st.set_page_config(page_title="進貨課", page_icon="🚚", layout="wide")
inject_logistics_theme()

ALLOW_PAGES = {
    "pages/1_驗收作業效能.py",
    "pages/2_上架作業效能.py",
    "pages/3_總揀作業效能.py",
    "pages/5_揀貨差異代庫存.py",
    "pages/27_QC未上架比對.py",
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
/* 條列式清單（跟你想要的一樣） */
.dept-list{ margin-top: 6px; }
.dept-row{
  display: flex;
  align-items: flex-start;
  gap: 10px;
  margin: 12px 0;
}
.dept-left{
  width: 34px;
  flex: 0 0 34px;
  display: inline-flex;
  align-items: flex-start;
  gap: 8px;
  margin-top: 2px;
}
.dept-bullet{
  color: rgba(15, 23, 42, 0.55);
  font-size: 16px;
  line-height: 1;
}
.dept-ico{
  font-size: 16px;
  line-height: 1;
}
.dept-right{
  flex: 1 1 auto;
  line-height: 1.55;
}
.dept-link{
  display: inline;
  color: rgba(15, 23, 42, 0.92) !important;
  font-weight: 900;
  font-size: 16px;
  line-height: 1.45;
  text-decoration: none !important;
  cursor: pointer;
}
.dept-link:hover{ opacity: 0.86; }
.dept-desc{
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
/* 同一視窗導頁 */
(function () {
  function bind() {
    document.querySelectorAll('a.dept-link').forEach(a => {
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
            f'<div class="dept-row">'
            f'  <div class="dept-left">'
            f'    <span class="dept-bullet">•</span>'
            f'    <span class="dept-ico">{icon}</span>'
            f'  </div>'
            f'  <div class="dept-right">'
            f'    <a class="dept-link" href="?page={encoded}" target="_self">{title}：</a>'
            f'    <span class="dept-desc">{desc}</span>'
            f'  </div>'
            f'</div>'
        ),
        unsafe_allow_html=True,
    )


def main():
    _route_by_query()

    set_page("進貨課", icon="🚚", subtitle="Inbound｜進貨相關模組入口")

    card_open("🚚 進貨課模組")
    _css_and_js()

    st.markdown('<div class="dept-list">', unsafe_allow_html=True)

    _nav_item(
        "✅",
        "驗收作業效能（KPI）",
        "pages/1_驗收作業效能.py",
        "人時效率、達標率、班別（AM/PM）切分、排除非作業區間（支援/離站/停機）",
    )
    _nav_item(
        "📦",
        "上架產能分析（Putaway KPI）",
        "pages/2_上架作業效能.py",
        "上架產能、人時效率、區塊/報表規則、班別切分",
    )
    _nav_item(
        "🎯",
        "總揀作業效能",
        "pages/3_總揀作業效能.py",
        "上午/下午達標分析、低空/高空門檻、排除非作業區間、匯出報表",
    )
    _nav_item(
        "🔎",
        "揀貨差異代庫存",
        "pages/5_揀貨差異代庫存.py",
        "少揀差異展開、庫存儲位/效期對應、國際條碼後五碼放大顯示",
    )
    _nav_item(
        "🧾",
        "QC 未上架比對",
        "pages/27_QC未上架比對.py",
        "QC「商品」比對「未上架明細商品碼」，回填「進貨日」，產生「符合未上架明細」，並刪除指定欄位",
    )
    st.markdown("</div>", unsafe_allow_html=True)
    card_close()


if __name__ == "__main__":
    main()
