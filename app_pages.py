class PageSpec:
    __slots__ = ("path", "title", "icon", "url_path", "default")

    def __init__(self, path: str, title: str, icon: str, url_path: str, default: bool = False):
        self.path = path
        self.title = title
        self.icon = icon
        self.url_path = url_path
        self.default = default


HIDDEN_NAV_LABELS = ["出貨課首頁", "進貨課首頁", "大樹KPI首頁", "大豐KPI首頁"]
HIDDEN_NAV_URL_KEYS = ["outbound-home", "inbound-home", "gt-kpi-home", "df-kpi-home"]


NAV_SECTIONS: dict[str, list[PageSpec]] = {
    "": [
        PageSpec("pages/0_首頁.py", "首頁", "🏠", "home", default=True),
        PageSpec("pages/31_資料庫後台.py", "資料庫後台", "🗄️", "database-admin"),
    ],
    "出貨課": [
        PageSpec("pages/7_出貨課首頁.py", "出貨課首頁", "📦", "outbound-home"),
        PageSpec("pages/6_撥貨差異.py", "撥貨差異", "📦", "outbound-transfer-diff"),
        PageSpec("pages/23_採品門市差異量.py", "採品門市差異量", "📄", "outbound-vendor-store-diff-23"),
        PageSpec("pages/24_出貨作業線產能.py", "出貨作業線產能", "📦", "outbound-line-productivity-24"),
        PageSpec("pages/29_各時段作業效率.py", "各時段作業效率", "⏱️", "outbound-hourly-efficiency-29"),
        PageSpec("pages/30_客訂差異.py", "客訂差異", "🧾", "outbound-custom-order-diff-30"),
    ],
    "進貨課": [
        PageSpec("pages/8_進貨課首頁.py", "進貨課首頁", "🚚", "inbound-home"),
        PageSpec("pages/1_驗收作業效能.py", "驗收作業效能", "✅", "inbound-qc"),
        PageSpec("pages/2_上架作業效能.py", "上架作業效能", "📦", "inbound-putaway"),
        PageSpec("pages/3_總揀作業效能.py", "總揀作業效能", "🎯", "inbound-pick"),
        PageSpec("pages/5_揀貨差異代庫存.py", "揀貨差異代庫存", "🔎", "inbound-pick-diff"),
        PageSpec("pages/27_QC未上架比對.py", "QC 未上架比對", "🧾", "inbound-qc-unputaway-compare-27"),
    ],
    "大樹KPI": [
        PageSpec("pages/9_大樹KPI首頁.py", "大樹KPI首頁", "📈", "gt-kpi-home"),
        PageSpec("pages/10_進貨驗收量.py", "進貨驗收量", "📥", "gt-inbound-receipt"),
        PageSpec("pages/11_庫存訂單應出量分析.py", "庫存訂單應出量分析", "📦", "gt-ship-should"),
        PageSpec("pages/12_越庫訂單分析.py", "越庫訂單分析", "🧾", "gt-xdock"),
        PageSpec("pages/13_庫存訂單實出量分析.py", "庫存訂單實出量分析", "🚚", "gt-ship-actual"),
        PageSpec("pages/14_每日上架分析.py", "每日上架分析", "📦", "gt-putaway-daily"),
        PageSpec("pages/15_庫存盤點正確率.py", "庫存盤點正確率", "🎯", "gt-inv-accuracy"),
        PageSpec("pages/16_門市到貨異常率.py", "門市到貨異常率", "🏪", "gt-store-arrival-abn"),
        PageSpec("pages/17_每日出勤工時分析.py", "每日出勤工時分析", "🕒", "gt-daily-attendance"),
        PageSpec("pages/18_各類儲區使用率.py", "各類儲區使用率", "🧊", "slot-zone-util-18"),
    ],
    "大豐KPI": [
        PageSpec("pages/19_大豐KPI首頁.py", "大豐KPI首頁", "📊", "df-kpi-home"),
        PageSpec("pages/20_進貨課 - 驗收量體.py", "進貨課 - 驗收量體", "✅", "df-qc-volume"),
        PageSpec("pages/21_進貨課 - 上架量體.py", "進貨課 - 上架量體", "📦", "df-putaway-volume"),
        PageSpec("pages/4_儲位使用率.py", "儲位使用率", "🧊", "inbound-slot-util"),
        PageSpec("pages/22_進貨課 - 總揀筆數.py", "進貨課 - 總揀筆數", "🎯", "df-pick-volume"),
        PageSpec("pages/28_每日庫存應作量.py", "每日庫存應作量", "🧮", "df-daily-inv-should-work-28"),
        PageSpec("pages/25_整體作業工時.py", "整體作業工時", "🕒", "df-total-workhours-25"),
        PageSpec("pages/26_整體作業量體.py", "整體作業量體", "🧹", "df-sort-volume-26"),
    ],
}
