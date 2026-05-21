# app.py
import os
import ast
import streamlit as st

st.set_page_config(
    page_title="大豐物流 - 作業平台",
    page_icon="assets/gf_logo.png",
    layout="wide",
)

# =========================
# Sidebar CSS + JS
# =========================
st.markdown(
    r"""
<style>
section[data-testid="stSidebar"]{
  padding-top: 10px;
  font-family: ui-sans-serif, system-ui, -apple-system, "Segoe UI",
               "Noto Sans TC", "Microsoft JhengHei", Arial, sans-serif;
}
section[data-testid="stSidebar"] a[data-testid="stSidebarNavLink"]{
  text-decoration: none !important;
}
section[data-testid="stSidebar"] a[data-testid="stSidebarNavLink"] *{
  font-size: 15.5px !important;
  font-weight: 650 !important;
  line-height: 1.55 !important;
  letter-spacing: .2px !important;
}
section[data-testid="stSidebar"] li a[data-testid="stSidebarNavLink"]{
  padding-top: 6px !important;
  padding-bottom: 6px !important;
}
section[data-testid="stSidebar"] ul > li:first-child a[data-testid="stSidebarNavLink"]{
  display:flex !important;
  align-items:center !important;
  justify-content:flex-start !important;
  gap:8px !important;
  padding: 10px 12px !important;
  min-height:48px !important;
  border-radius: 12px !important;
}
section[data-testid="stSidebar"] ul > li:first-child a[data-testid="stSidebarNavLink"] *{
  font-size: 26px !important;
  font-weight: 900 !important;
  line-height: 1.15 !important;
  white-space: nowrap !important;
  text-align: left !important;
  letter-spacing: .3px !important;
}
section[data-testid="stSidebar"] [data-testid="stSidebarNav"] h2,
section[data-testid="stSidebar"] [data-testid="stSidebarNav"] h3,
section[data-testid="stSidebar"] [data-testid="stSidebarNav"] h4{
  font-size: 13.5px !important;
  font-weight: 850 !important;
  color: rgba(15,23,42,.72) !important;
  letter-spacing: .9px !important;
  margin: 14px 0 6px !important;
}
section[data-testid="stSidebar"] a[data-testid="stSidebarNavLink"]{ gap: 8px !important; }

/* 隱藏群組首頁 */
section[data-testid="stSidebar"] li:has(a[data-testid="stSidebarNavLink"][href*="outbound-home"]){ display:none !important; }
section[data-testid="stSidebar"] li:has(a[data-testid="stSidebarNavLink"][href*="inbound-home"]){  display:none !important; }
section[data-testid="stSidebar"] li:has(a[data-testid="stSidebarNavLink"][href*="gt-kpi-home"]){    display:none !important; }
section[data-testid="stSidebar"] li:has(a[data-testid="stSidebarNavLink"][href*="df-kpi-home"]){    display:none !important; }

section[data-testid="stSidebar"] li:has(span[label="出貨課首頁"]){ display:none !important; }
section[data-testid="stSidebar"] li:has(span[label="進貨課首頁"]){ display:none !important; }
section[data-testid="stSidebar"] li:has(span[label="大樹KPI首頁"]){ display:none !important; }
section[data-testid="stSidebar"] li:has(span[label="大豐KPI首頁"]){ display:none !important; }
</style>

<script>
(function () {
  const HIDE_LABELS = ["出貨課首頁", "進貨課首頁", "大樹KPI首頁", "大豐KPI首頁"];
  const HIDE_KEYS   = ["outbound-home", "inbound-home", "gt-kpi-home", "df-kpi-home"];
  
  function hideByHrefAndLabel(){
    const sidebar = document.querySelector('section[data-testid="stSidebar"]');
    if(!sidebar) return;
    const links = sidebar.querySelectorAll('a[data-testid="stSidebarNavLink"]');

    links.forEach(a => {
      const href = (a.getAttribute("href") || a.href || "");
      const labelSpan = a.querySelector('span[label]');
      const label = labelSpan ? (labelSpan.getAttribute("label") || "") : "";
      const hitHref  = HIDE_KEYS.some(k => href.includes(k));
      const hitLabel = HIDE_LABELS.includes(label);
      if(hitHref || hitLabel){
        const li = a.closest("li");
        if(li) li.style.display = "none";
        a.style.display = "none";
      }
    });
  }

  const root = document.querySelector('#root') || document.body;
  const obs = new MutationObserver(() => hideByHrefAndLabel());
  obs.observe(root, { childList: true, subtree: true });

  hideByHrefAndLabel();
  setTimeout(hideByHrefAndLabel, 50);
  setTimeout(hideByHrefAndLabel, 200);
  setTimeout(hideByHrefAndLabel, 800);
  setTimeout(hideByHrefAndLabel, 2000);
})();
</script>
""",
    unsafe_allow_html=True,
)

# =========================
# ✅ Preflight: 語法/縮排檢查（避免某頁壞掉整站掛）
# =========================
BROKEN_PAGES: list[tuple[str, str]] = []
MISSING_PAGES: list[str] = []


def _syntax_ok(path: str) -> bool:
    try:
        with open(path, "r", encoding="utf-8") as f:
            src = f.read()
        ast.parse(src, filename=path)
        return True
    except Exception as e:
        BROKEN_PAGES.append((path, repr(e)))
        return False


def page_if_exists(path: str, title: str, icon: str, **kwargs):
    if not os.path.exists(path):
        MISSING_PAGES.append(path)
        return None
    if not _syntax_ok(path):
        return None
    try:
        return st.Page(path, title=title, icon=icon, **kwargs)
    except Exception as e:
        BROKEN_PAGES.append((path, f"st.Page 建立失敗：{repr(e)}"))
        return None


# =========================
# ✅ Pages
# =========================
home_page = page_if_exists(
    "pages/0_首頁.py",
    "首頁",
    "🏠",
    default=True,
    url_path="home",
)

# =========================
# 出貨課
# =========================
outbound_home = page_if_exists(
    "pages/7_出貨課首頁.py",
    "出貨課首頁",
    "📦",
    url_path="outbound-home",
)

transfer_diff_page = page_if_exists(
    "pages/6_撥貨差異.py",
    "撥貨差異",
    "📦",
    url_path="outbound-transfer-diff",
)

outbound_vendor_store_diff = page_if_exists(
    "pages/23_採品門市差異量.py",
    "採品門市差異量",
    "📄",
    url_path="outbound-vendor-store-diff-23",
)

outbound_line_productivity = page_if_exists(
    "pages/24_出貨作業線產能.py",
    "出貨作業線產能",
    "📦",
    url_path="outbound-line-productivity-24",
)

outbound_hourly_efficiency = page_if_exists(
    "pages/29_各時段作業效率.py",
    "各時段作業效率",
    "⏱️",
    url_path="outbound-hourly-efficiency-29",
)

outbound_custom_order_diff = page_if_exists(
    "pages/30_客訂差異.py",
    "客訂差異",
    "🧾",
    url_path="outbound-custom-order-diff-30",
)

# =========================
# 進貨課
# =========================
inbound_home = page_if_exists(
    "pages/8_進貨課首頁.py",
    "進貨課首頁",
    "🚚",
    url_path="inbound-home",
)

qc_page = page_if_exists(
    "pages/1_驗收作業效能.py",
    "驗收作業效能",
    "✅",
    url_path="inbound-qc",
)

putaway_page = page_if_exists(
    "pages/2_上架作業效能.py",
    "上架作業效能",
    "📦",
    url_path="inbound-putaway",
)

pick_page = page_if_exists(
    "pages/3_總揀作業效能.py",
    "總揀作業效能",
    "🎯",
    url_path="inbound-pick",
)

diff_page = page_if_exists(
    "pages/5_揀貨差異代庫存.py",
    "揀貨差異代庫存",
    "🔎",
    url_path="inbound-pick-diff",
)

qc_unputaway_compare = page_if_exists(
    "pages/27_QC未上架比對.py",
    "QC 未上架比對",
    "🧾",
    url_path="inbound-qc-unputaway-compare-27",
)

# =========================
# 大樹KPI
# =========================
gt_kpi_home = page_if_exists(
    "pages/9_大樹KPI首頁.py",
    "大樹KPI首頁",
    "📈",
    url_path="gt-kpi-home",
)

gt_inbound_receipt = page_if_exists(
    "pages/10_進貨驗收量.py",
    "進貨驗收量",
    "📥",
    url_path="gt-inbound-receipt",
)

gt_ship_should = page_if_exists(
    "pages/11_庫存訂單應出量分析.py",
    "庫存訂單應出量分析",
    "📦",
    url_path="gt-ship-should",
)

gt_xdock = page_if_exists(
    "pages/12_越庫訂單分析.py",
    "越庫訂單分析",
    "🧾",
    url_path="gt-xdock",
)

gt_ship_actual = page_if_exists(
    "pages/13_庫存訂單實出量分析.py",
    "庫存訂單實出量分析",
    "🚚",
    url_path="gt-ship-actual",
)

gt_putaway_daily = page_if_exists(
    "pages/14_每日上架分析.py",
    "每日上架分析",
    "📦",
    url_path="gt-putaway-daily",
)

gt_inv_accuracy = page_if_exists(
    "pages/15_庫存盤點正確率.py",
    "庫存盤點正確率",
    "🎯",
    url_path="gt-inv-accuracy",
)

gt_store_arrival_abn = page_if_exists(
    "pages/16_門市到貨異常率.py",
    "門市到貨異常率",
    "🏪",
    url_path="gt-store-arrival-abn",
)

gt_daily_attendance = page_if_exists(
    "pages/17_每日出勤工時分析.py",
    "每日出勤工時分析",
    "🕒",
    url_path="gt-daily-attendance",
)

slot_util_page = page_if_exists(
    "pages/18_各類儲區使用率.py",
    "各類儲區使用率",
    "🧊",
    url_path="slot-zone-util-18",
)

# =========================
# 大豐KPI
# =========================
df_kpi_home = page_if_exists(
    "pages/19_大豐KPI首頁.py",
    "大豐KPI首頁",
    "📊",
    url_path="df-kpi-home",
)

df_qc_volume = page_if_exists(
    "pages/20_進貨課 - 驗收量體.py",
    "進貨課 - 驗收量體",
    "✅",
    url_path="df-qc-volume",
)

df_putaway_volume = page_if_exists(
    "pages/21_進貨課 - 上架量體.py",
    "進貨課 - 上架量體",
    "📦",
    url_path="df-putaway-volume",
)

slot_page = page_if_exists(
    "pages/4_儲位使用率.py",
    "儲位使用率",
    "🧊",
    url_path="inbound-slot-util",
)

df_pick_volume = page_if_exists(
    "pages/22_進貨課 - 總揀筆數.py",
    "進貨課 - 總揀筆數",
    "🎯",
    url_path="df-pick-volume",
)

df_daily_inv_should_work = page_if_exists(
    "pages/28_每日庫存應作量.py",
    "每日庫存應作量",
    "🧮",
    url_path="df-daily-inv-should-work-28",
)

df_total_workhours = page_if_exists(
    "pages/25_整體作業工時.py",
    "整體作業工時",
    "🕒",
    url_path="df-total-workhours-25",
)

df_sort_volume = page_if_exists(
    "pages/26_整體作業量體.py",
    "整體作業量體",
    "🧹",
    url_path="df-sort-volume-26",
)

# =========================
# ✅ Sidebar 顯示「壞頁 / 缺檔」清單
# =========================
if MISSING_PAGES:
    with st.sidebar.expander("⚠️ 找不到檔案（未載入）", expanded=False):
        st.caption("下列 pages 檔案不存在，所以不會出現在左側選單：")
        for p in MISSING_PAGES:
            st.code(p)

if BROKEN_PAGES:
    with st.sidebar.expander("⚠️ 已停用頁面（語法/縮排錯）", expanded=True):
        st.caption("以下檔案有 IndentationError / SyntaxError，已自動略過避免整站掛掉：")
        for p, err in BROKEN_PAGES:
            st.code(f"{p}\n{err}")

# =========================
# ✅ Navigation
# =========================
pg = st.navigation(
    {
        "": [
            p
            for p in [
                home_page,
            ]
            if p
        ],
        "出貨課": [
            p
            for p in [
                outbound_home,
                transfer_diff_page,
                outbound_vendor_store_diff,
                outbound_line_productivity,
                outbound_hourly_efficiency,
                outbound_custom_order_diff,
            ]
            if p
        ],
        "進貨課": [
            p
            for p in [
                inbound_home,
                qc_page,
                putaway_page,
                pick_page,
                diff_page,
                qc_unputaway_compare,
            ]
            if p
        ],
        "大樹KPI": [
            p
            for p in [
                gt_kpi_home,
                gt_inbound_receipt,
                gt_ship_should,
                gt_xdock,
                gt_ship_actual,
                gt_putaway_daily,
                gt_inv_accuracy,
                gt_store_arrival_abn,
                gt_daily_attendance,
                slot_util_page,
            ]
            if p
        ],
        "大豐KPI": [
            p
            for p in [
                df_kpi_home,
                df_qc_volume,
                df_putaway_volume,
                slot_page,
                df_pick_volume,
                df_daily_inv_should_work,
                df_total_workhours,
                df_sort_volume,
            ]
            if p
        ],
    },
    expanded=False,
)

pg.run()
