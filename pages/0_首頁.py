# pages/0_首頁.py
import streamlit as st
from urllib.parse import quote, unquote

from common_ui import inject_logistics_theme, set_page, card_open, card_close

st.set_page_config(page_title="大豐物流 - 作業平台", page_icon="🚚", layout="wide")
inject_logistics_theme()

# ✅ 只允許導去這三個入口（避免 switch_page 找不到而整頁空白）
ALLOW = {
    "pages/7_出貨課首頁.py",
    "pages/8_進貨課首頁.py",
    "pages/9_大樹KPI首頁.py",
    "pages/19_大豐KPI首頁.py", 
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

    if target not in ALLOW:
        # 不要讓錯誤路徑把首頁打掛
        return

    try:
        st.switch_page(target)
    except Exception:
        return

def _css_and_js():
    st.markdown(
        r"""
<style>
/* 入口方框（並排） */
.entry-grid{
  display:grid;
  grid-template-columns: repeat(3, minmax(220px, 1fr));
  gap: 16px;
  margin-top: 12px;
}
@media (max-width: 1100px){ .entry-grid{ grid-template-columns: repeat(2, minmax(220px, 1fr)); } }
@media (max-width: 700px){ .entry-grid{ grid-template-columns: repeat(1, minmax(220px, 1fr)); } }

.entry{
  position: relative;
  border-radius: 8px;
  border: 1px solid #D8DEE8;
  background: #FFFFFF;
  padding: 16px 18px 14px;
  min-height: 104px;
  box-shadow: 0 1px 2px rgba(15,23,42,0.05);
  overflow: hidden;
  transition: border-color .12s ease, box-shadow .12s ease;
}
.entry:hover{
  box-shadow: 0 4px 12px rgba(15,23,42,0.08);
  border-color: #0B3A67;
}
.entry-head{ display:flex; align-items:center; gap:10px; }
.entry-ico{
  width: 36px; height: 36px;
  border-radius: 6px;
  display:flex; align-items:center; justify-content:center;
  font-size: 18px;
  border: 1px solid #D8DEE8;
  background: #F4F7FB;
}
.entry-name{ font-size:16px; font-weight:800; line-height:1.2; color: #0F172A; }
.entry-desc{ margin-top:8px; font-size:12.5px; font-weight:650; color: #5B667A; line-height:1.45; padding-right: 56px; }
.entry-cta{
  position:absolute; right: 12px; bottom: 10px;
  font-size: 12px; font-weight: 800;
  color: #0B3A67;
}

a.entry-link{ text-decoration:none !important; color: inherit !important; display:block; }
div[data-testid="stMarkdown"]{ margin: 0 !important; }
</style>

<script>
(function () {
  function bind() {
    document.querySelectorAll('a.entry-link').forEach(a => {
      if (a.dataset.bound === "1") return;
      a.dataset.bound = "1";
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

def _tile(icon: str, title: str, desc: str, page_path: str) -> str:
    encoded = quote(page_path, safe="/_.-")
    return (
        f'<a class="entry-link" href="?page={encoded}" target="_self">'
        f'  <div class="entry">'
        f'    <div class="entry-head">'
        f'      <div class="entry-ico">{icon}</div>'
        f'      <div class="entry-name">{title}</div>'
        f'    </div>'
        f'    <div class="entry-desc">{desc}</div>'
        f'    <div class="entry-cta">進入 →</div>'
        f'  </div>'
        f'</a>'
    )

def main():
    _route_by_query()

    set_page(
        "大豐物流 - 作業平台",
        icon="🚚",
        subtitle="作業KPI｜班別分析（AM/PM）｜排除非作業區間",
    )

    card_open("📌 課別入口")
    _css_and_js()

    tiles = [
        _tile("📦", "出貨課", "撥貨差異｜出貨/包裝/異常（進入後以條列式顯示）", "pages/7_出貨課首頁.py"),
        _tile("🚚", "進貨課", "驗收/上架/總揀/儲位/差異代庫存（進入後以條列式顯示）", "pages/8_進貨課首頁.py"),
        _tile("📈", "大樹KPI", "KPI 模組入口｜匯總｜趨勢（進入後以條列式顯示）", "pages/9_大樹KPI首頁.py"),
        _tile("📊", "大豐KPI", "KPI 模組入口｜進貨課整體｜出貨課整體（進入後以條列式顯示）", "pages/19_大豐KPI首頁.py"),
    ]

    st.markdown('<div class="entry-grid">' + "".join(tiles) + "</div>", unsafe_allow_html=True)
    card_close()

if __name__ == "__main__":
    main()
