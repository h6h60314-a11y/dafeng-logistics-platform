# pages/19_大豐KPI首頁.py
import streamlit as st
from urllib.parse import quote, unquote

from common_ui import inject_logistics_theme, set_page, card_open, card_close

st.set_page_config(page_title="大豐KPI", page_icon="📊", layout="wide")
inject_logistics_theme()

# ✅ 允許從 KPI 首頁導頁的模組清單（安全白名單）
ALLOW_PAGES = {
    "pages/20_進貨課 - 驗收量體.py",
    "pages/21_進貨課 - 上架量體.py",
    "pages/4_儲位使用率.py",
    "pages/22_進貨課 - 總揀筆數.py",
    "pages/28_每日庫存應作量.py",
    "pages/25_整體作業工時.py",
    "pages/26_整體作業量體.py",  
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

    set_page("大豐KPI", icon="📊", subtitle="KPI 模組入口｜匯總｜告警｜趨勢")

    card_open("📊 大豐KPI模組")
    _css_and_js()

    st.markdown('<div class="kpi-list">', unsafe_allow_html=True)

    _nav_item(
        "✅",
        "進貨課 - 驗收量體",
        "pages/20_進貨課 - 驗收量體.py",
        "只保留「到=QC」｜SKU（唯一商品）｜ITEM（筆數）｜輸出Excel",
    )

    _nav_item(
        "📦",
        "進貨課 - 上架量體",
        "pages/21_進貨課 - 上架量體.py",
        "由=QC｜到排除關鍵字｜對應儲位類型｜高低空統計｜輸出Excel",
    )

    _nav_item(
        "🧊",
        "儲位使用率分析",
        "pages/4_儲位使用率.py",
        "依區(溫層)分類統計、使用率門檻提示、分類可調整、KPI圖格呈現",
    )

    _nav_item(
        "🎯",
        "進貨課 - 總揀筆數",
        "pages/22_進貨課 - 總揀筆數.py",
        "多檔批次｜成箱/零散(或ALL)｜排除儲位｜回填儲位類型｜單頁Excel輸出",
    )

    _nav_item(
        "🧮",
        "每日庫存應作量",
        "pages/28_每日庫存應作量.py",
        "訂單檔＋商品主檔｜排除特殊儲位｜補碼｜大類加權×計量單位數量｜CSV下載",
    )

    _nav_item(
        "🕒",
        "整體作業工時",
        "pages/25_整體作業工時.py",
        "出勤報表｜排除空打卡＋外倉職務｜工時摘要＋明細下載",
    )

    _nav_item(
        "🧹",
        "整體作業量體",
        "pages/26_整體作業量體.py",
        "刪除箱類型含站所｜計量單位數量＋出貨單位｜GM/一般倉統計｜Excel下載",
    )

    st.markdown("</div>", unsafe_allow_html=True)
    card_close()


if __name__ == "__main__":
    main()
