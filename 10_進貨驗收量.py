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
  gap: 14px;
  margin-top: 10px;
}
@media (max-width: 1100px){ .entry-grid{ grid-template-columns: repeat(2, minmax(220px, 1fr)); } }
@media (max-width: 700px){ .entry-grid{ grid-template-columns: repeat(1, minmax(220px, 1fr)); } }

.entry{
  position: relative;
  border-radius: 16px;
  border: 1px solid rgba(15,23,42,0.10);
  background: rgba(255,255,255,0.92);
  padding: 14px 14px 12px;
  min-height: 96px;
  box-shadow: 0 14px 26px rgba(2,6,23,0.06);
  overflow: hidden;
  transition: transform .08s ease, box-shadow .12s ease, border-color .12s ease;
}
.entry:hover{
  transform: translateY(-1px);
  box-shadow: 0 18px 32px rgba(2,6,23,0.10);
  border-color: rgba(15,23,42,0.18);
}
.entry-head{ display:flex; align-items:center; gap:10px; }
.entry-ico{
  width: 34px; height: 34px;
  border-radius: 12px;
  display:flex; align-items:center; justify-content:center;
  font-size: 18px;
  border: 1px solid rgba(15,23,42,0.10);
  background: rgba(255,255,255,0.85);
}
.entry-name{ font-size:16px; font-weight:950; line-height:1.15; color: rgba(15,23,42,0.92); }
.entry-desc{ margin-top:6px; font-size:12px; font-weight:850; color: rgba(15,23,42,0.62); line-height:1.35; }
.entry-cta{
  position:absolute; right: 12px; bottom: 10px;
  font-size: 12px; font-weight: 950;
  color: rgba(15,23,42,0.55);
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
