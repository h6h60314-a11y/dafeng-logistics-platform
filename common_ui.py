from __future__ import annotations

import io
from dataclasses import dataclass
from typing import List, Optional, Sequence, Dict, Any

import pandas as pd
import streamlit as st


# =========================================================
# Theme / CSS（物流專業風格）
# =========================================================
def inject_logistics_theme():
    """
    Logistics / Warehouse dashboard style.
    - 中間內容區更寬（完整呈現）
    - 全站字體縮小一點（含側欄、Metric）
    - 下載按鈕一致化
    """
    st.markdown(
        """
<style>
:root{
  --ink: rgba(15, 23, 42, 0.92);
  --muted: rgba(15, 23, 42, 0.60);
  --line: rgba(15, 23, 42, 0.10);
  --card: rgba(255,255,255,0.88);
  --card2: rgba(255,255,255,0.70);
  --blue: rgba(2, 132, 199, 1.00);
  --blueSoft: rgba(2, 132, 199, 0.12);
  --blueSoft2: rgba(2, 132, 199, 0.18);
  --badBg: #FDE2E2;
  --badText: #7F1D1D;
}

.stApp {
  color: var(--ink);
  background: radial-gradient(1200px 700px at 20% 0%, rgba(2,132,199,0.10) 0%, #f5f8fc 55%, #ecf2fa 100%);
}

/* remove top bar feeling */
header[data-testid="stHeader"] { background: transparent !important; }
div[data-testid="stDecoration"] { display: none; }

/* ============== Layout: make center wider ============== */
.block-container{
  max-width: 1600px !important;
  padding-top: 0.8rem !important;
  padding-bottom: 1.6rem !important;
  padding-left: 1.0rem !important;
  padding-right: 1.0rem !important;
}
@media (min-width: 1800px){
  .block-container{ max-width: 1800px !important; }
}

/* ============== Global font scale (smaller) ============== */
html, body, [class*="st-"], .stApp{
  font-size: 14px !important;
}

/* 標題縮小 */
h1 { font-size: 28px !important; }
h2 { font-size: 22px !important; }
h3 { font-size: 16px !important; }

/* Sidebar */
section[data-testid="stSidebar"]{
  background: #f8fafc;
  border-right: 1px solid var(--line);
}
section[data-testid="stSidebar"] *{
  font-size: 13px !important;
}

/* Card */
._gt_card{
  border: 1px solid var(--line);
  background: var(--card);
  border-radius: 20px;
  padding: 14px 14px 10px 14px;
  margin-bottom: 12px;
  box-shadow: 0 10px 30px rgba(15,23,42,0.06);
}
._gt_card h3{
  margin: 0 0 8px 0;
  letter-spacing: .2px;
}

/* Helpers */
._gt_hint{
  color: var(--muted);
  font-size: 13px;
  line-height: 1.55;
  font-weight: 650;
}
._gt_badge{
  display:inline-block;
  padding:2px 10px;
  border-radius:999px;
  border:1px solid var(--line);
  font-size:12px;
  font-weight:800;
  background:#fff;
}

/* Tables */
div[data-testid="stDataFrame"]{
  border-radius: 16px;
  overflow: hidden;
  border: 1px solid var(--line);
  background: var(--card2);
}

/* KPI Metric size */
[data-testid="stMetricLabel"]{
  font-size: 12px !important;
  color: var(--muted) !important;
  font-weight: 700 !important;
}
[data-testid="stMetricValue"]{
  font-size: 22px !important;
  font-weight: 900 !important;
}

/* Buttons */
.stButton > button{
  border-radius: 14px;
  border: 1px solid rgba(2, 132, 199, 0.30);
  background: var(--blueSoft);
  color: var(--ink);
  padding: 0.50rem 0.85rem;
  font-weight: 850;
}
.stButton > button:hover{
  border: 1px solid rgba(2, 132, 199, 0.45);
  background: var(--blueSoft2);
}

/* Download button */
div[data-testid="stDownloadButton"] button{
  border-radius: 14px !important;
  border: 1px solid rgba(2, 132, 199, 0.30) !important;
  background: var(--blueSoft) !important;
  color: var(--ink) !important;
  padding: 0.60rem 0.95rem !important;
  font-weight: 900 !important;
}
div[data-testid="stDownloadButton"] button:hover{
  border: 1px solid rgba(2, 132, 199, 0.45) !important;
  background: var(--blueSoft2) !important;
}

/* Uploader */
div[data-testid="stFileUploaderDropzone"]{
  border-radius: 18px;
  border: 1px dashed rgba(15, 23, 42, 0.22);
  background: rgba(255,255,255,0.80);
}
</style>
""",
        unsafe_allow_html=True,
    )


# Backward compatibility
def inject_purple_theme():
    inject_logistics_theme()


# =========================================================
# Page helpers
# =========================================================
def set_page(title: str, icon: str = "🏭", subtitle: Optional[str] = None):
    inject_logistics_theme()
    st.markdown(f"## {icon} {title}")
    if subtitle:
        st.markdown(f'<div class="_gt_hint">{subtitle}</div>', unsafe_allow_html=True)


def hint(text: str):
    st.markdown(f'<div class="_gt_hint">{text}</div>', unsafe_allow_html=True)


def badge(text: str):
    st.markdown(f'<span class="_gt_badge">{text}</span>', unsafe_allow_html=True)


def card_open(title: str, right_badge: Optional[str] = None):
    """
    有標題卡片（會產生 <h3>）
    """
    if right_badge:
        st.markdown(
            f'<div class="_gt_card">'
            f'<div style="display:flex;justify-content:space-between;align-items:center;gap:10px;">'
            f'<h3 style="margin:0;">{title}</h3>'
            f'<span class="_gt_badge">{right_badge}</span>'
            f'</div>',
            unsafe_allow_html=True,
        )
    else:
        st.markdown(f'<div class="_gt_card"><h3>{title}</h3>', unsafe_allow_html=True)


def card_open_plain():
    """
    無標題卡片（不會產生 <h3>，用於「按鈕=文字」避免分段）
    """
    st.markdown('<div class="_gt_card">', unsafe_allow_html=True)


def card_close():
    st.markdown("</div>", unsafe_allow_html=True)


# =========================================================
# KPI / Metrics
# =========================================================
@dataclass
class KPI:
    label: str
    value: str
    delta: Optional[str] = None


def render_kpis(kpis: Sequence[KPI], cols: Optional[int] = None):
    if not kpis:
        return
    n = cols or min(5, len(kpis))
    columns = st.columns(n)
    for i, k in enumerate(kpis):
        with columns[i % n]:
            if k.delta is None:
                st.metric(label=k.label, value=k.value)
            else:
                st.metric(label=k.label, value=k.value, delta=k.delta)


# =========================================================
# KPI Table Styling（低於門檻 → 紅色）
# =========================================================
def style_kpi_below_target(df: pd.DataFrame, eff_col: str, target: float):
    def _row_style(row):
        try:
            val = float(row.get(eff_col))
        except Exception:
            return [""] * len(row)

        if val < target:
            return ["background-color: #FDE2E2; color: #7F1D1D; font-weight: 650"] * len(row)
        return [""] * len(row)

    return df.style.apply(_row_style, axis=1)


def show_kpi_table(df: pd.DataFrame, *, eff_col: str, target: float):
    if df is None or df.empty:
        st.info("目前沒有可顯示的資料")
        return

    st.dataframe(
        style_kpi_below_target(df, eff_col=eff_col, target=target),
        use_container_width=True,
        hide_index=True,
    )


# =========================================================
# Charts（TopN：低於 target 紅色）
# =========================================================
def bar_topN(
    df: pd.DataFrame,
    x_col: str,
    y_col: str,
    hover_cols: Optional[List[str]] = None,
    top_n: int = 30,
    target: Optional[float] = None,
    title: str = "",
):
    if df is None or df.empty:
        st.info("目前無資料可視覺化")
        return

    data = df.copy()
    keep_cols = [c for c in [x_col, y_col] + (hover_cols or []) if c in data.columns]
    data = data[keep_cols].copy()

    data[y_col] = pd.to_numeric(data[y_col], errors="coerce")
    data = data.dropna(subset=[y_col])
    data = data.sort_values(y_col, ascending=False).head(int(top_n))

    try:
        import altair as alt  # type: ignore

        if target is not None:
            color_enc = alt.condition(
                alt.datum[y_col] < float(target),
                alt.value("#DC2626"),
                alt.value("#0284C7"),
            )
        else:
            color_enc = alt.value("#0284C7")

        base = (
            alt.Chart(data)
            .mark_bar()
            .encode(
                x=alt.X(f"{y_col}:Q", title=y_col),
                y=alt.Y(f"{x_col}:N", sort="-x", title=""),
                color=color_enc,
                tooltip=[c for c in [x_col, y_col] + (hover_cols or []) if c in data.columns],
            )
            .properties(height=min(560, 28 * max(6, len(data))))
        )

        layers = [base]
        if target is not None:
            rule = (
                alt.Chart(pd.DataFrame({"target": [float(target)]}))
                .mark_rule(strokeDash=[6, 4])
                .encode(x="target:Q")
            )
            layers.append(rule)

        if title:
            st.caption(title)

        st.altair_chart(alt.layer(*layers), use_container_width=True)

    except Exception:
        st.bar_chart(data.set_index(x_col)[y_col])


# =========================================================
# Sidebar Controls（排除區間：手動輸入 HH:MM）
# =========================================================
@dataclass
class ExcludeWindow:
    start: str  # HH:MM
    end: str    # HH:MM
    data_entry: str = ""


def _init_exclude_state(key: str):
    if key not in st.session_state:
        st.session_state[key] = []


def sidebar_controls(
    *,
    default_top_n: int = 30,
    enable_exclude_windows: bool = True,
    state_key_prefix: str = "gt",
) -> Dict[str, Any]:
    inject_logistics_theme()
    result: Dict[str, Any] = {}

    st.sidebar.markdown("### ⚙️ 計算條件設定")
    st.sidebar.markdown('<div class="_gt_hint">本區設定僅影響本次分析</div>', unsafe_allow_html=True)

    top_n = st.sidebar.number_input(
        "效率排行顯示人數（Top N）",
        min_value=5,
        max_value=200,
        value=int(default_top_n),
        step=1,
    )
    result["top_n"] = int(top_n)

    if enable_exclude_windows:
        st.sidebar.markdown("---")
        st.sidebar.markdown("### ⛔ 排除區間（非作業時段）")
        st.sidebar.markdown('<div class="_gt_hint">請手動輸入時間（HH:MM），例如 12:30</div>', unsafe_allow_html=True)

        state_key = f"{state_key_prefix}_exclude_windows"
        _init_exclude_state(state_key)

        with st.sidebar.expander("新增排除區間", expanded=False):
            data_entry = st.text_input("資料登錄人（可留空）", value="", key=f"{state_key_prefix}_ex_data_entry")

            c1, c2 = st.columns(2)
            with c1:
                start_str = st.text_input(
                    "開始時間（HH:MM）",
                    value="08:00",
                    placeholder="例如 08:00",
                    key=f"{state_key_prefix}_ex_start_str",
                )
            with c2:
                end_str = st.text_input(
                    "結束時間（HH:MM）",
                    value="08:30",
                    placeholder="例如 12:30",
                    key=f"{state_key_prefix}_ex_end_str",
                )

            if st.button("＋ 新增排除區間", key=f"{state_key_prefix}_btn_add_ex"):
                try:
                    s_time = pd.to_datetime(start_str, format="%H:%M").time()
                    e_time = pd.to_datetime(end_str, format="%H:%M").time()

                    if s_time >= e_time:
                        st.error("❌ 開始時間需早於結束時間")
                    else:
                        st.session_state[state_key].append(
                            ExcludeWindow(
                                start=start_str,
                                end=end_str,
                                data_entry=(data_entry or "").strip(),
                            )
                        )
                        st.success(f"已新增：{start_str} - {end_str}")
                except Exception:
                    st.error("❌ 時間格式錯誤，請使用 HH:MM（例如 08:00）")

        windows: List[ExcludeWindow] = st.session_state[state_key]
        if windows:
            st.sidebar.markdown("#### 已設定排除區間")
            for idx, w in enumerate(list(windows)):
                cols = st.sidebar.columns([0.72, 0.28])
                with cols[0]:
                    label = f"{w.start} - {w.end}"
                    if (w.data_entry or "").strip():
                        label += f"｜登錄：{w.data_entry}"
                    st.write(label)
                with cols[1]:
                    if st.button("刪除", key=f"{state_key_prefix}_ex_del_{idx}"):
                        st.session_state[state_key].pop(idx)
                        st.experimental_rerun()
        else:
            st.sidebar.info("尚未新增排除區間")

        result["exclude_windows"] = [
            {"start": w.start, "end": w.end, "data_entry": w.data_entry} for w in st.session_state[state_key]
        ]
    else:
        result["exclude_windows"] = []

    return result


# =========================================================
# Downloads
# =========================================================
def download_excel(
    xlsx_bytes: bytes,
    filename: str = "KPI報表.xlsx",
    label: str = "⬇️ 匯出 KPI 報表",
    use_container_width: bool = True,
):
    st.download_button(
        label=label,
        data=xlsx_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=use_container_width,
    )


def download_excel_button(
    xlsx_bytes: bytes,
    filename: str = "KPI報表.xlsx",
    label: str = "⬇️ 匯出 KPI 報表",
):
    """
    你要的效果：畫面只看到「一行」，而且那一行就是按鈕。
    用法：頁面上直接呼叫（不要再 card_open）
    """
    st.download_button(
        label=label,
        data=xlsx_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )


def download_excel_card(
    xlsx_bytes: bytes,
    filename: str = "KPI報表.xlsx",
    label: str = "⬇️ 匯出 KPI 報表",
):
    """
    卡片外框 + 內部按鈕（✅ 無標題，不會分成兩段）
    """
    card_open_plain()
    download_excel_button(xlsx_bytes=xlsx_bytes, filename=filename, label=label)
    card_close()


# =========================================================
# Excel helpers
# =========================================================
def dataframe_to_excel_bytes(sheets: Dict[str, pd.DataFrame]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for name, df in sheets.items():
            if df is None:
                continue
            safe_name = str(name)[:31] or "Sheet1"
            df.to_excel(writer, sheet_name=safe_name, index=False)
    return output.getvalue()


# =========================================================
# Optional: table_block
# =========================================================
def table_block(
    summary_title: str,
    summary_df: pd.DataFrame,
    detail_title: str = "",
    detail_df: Optional[pd.DataFrame] = None,
    detail_expanded: bool = False,
    *,
    style_eff_col: Optional[str] = None,
    style_target: Optional[float] = None,
):
    card_open(summary_title)
    if summary_df is None or summary_df.empty:
        st.info("目前沒有可顯示的資料")
    else:
        if style_eff_col is not None and style_target is not None and style_eff_col in summary_df.columns:
            st.dataframe(
                style_kpi_below_target(summary_df, eff_col=style_eff_col, target=float(style_target)),
                use_container_width=True,
                hide_index=True,
            )
        else:
            st.dataframe(summary_df, use_container_width=True, hide_index=True)
    card_close()

    if detail_title:
        with st.expander(detail_title, expanded=detail_expanded):
            if detail_df is None or detail_df.empty:
                st.info("目前沒有明細資料")
            else:
                st.dataframe(detail_df, use_container_width=True, hide_index=True)
