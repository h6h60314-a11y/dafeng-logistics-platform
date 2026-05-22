from __future__ import annotations

import streamlit as st
import pandas as pd

from common_ui import (
    set_page,               # set_page 內會 inject_logistics_theme()
    KPI,
    render_kpis,
    bar_topN,
    download_excel_card,    # ✅ 一行=按鈕、卡片外框、不分段
    sidebar_controls,       # ✅ 統一左側設定（TopN + 排除空窗）
    card_open,
    card_close,
    show_kpi_table,         # ✅ 整列紅/綠顯示（效率 < target 會紅）
)

from qc_core import run_qc_efficiency


# =========================================================
# Helpers
# =========================================================
def _adapt_exclude_windows_to_skip_rules(exclude_windows):
    """
    將 common_ui.sidebar_controls() 的 exclude_windows 格式：
      [{"start":"HH:MM","end":"HH:MM","data_entry":""}, ...]
    轉回 qc_core.run_qc_efficiency 需要的 skip_rules 格式：
      [{"user":"", "t_start": datetime.time, "t_end": datetime.time}, ...]
    """
    skip_rules = []
    for w in exclude_windows or []:
        start_str = (w.get("start") or "").strip()
        end_str = (w.get("end") or "").strip()
        user_str = (w.get("data_entry") or "").strip()

        try:
            # 明確指定 HH:MM，避免被解析成日期
            s = pd.to_datetime(start_str, format="%H:%M").time()
            e = pd.to_datetime(end_str, format="%H:%M").time()
        except Exception:
            continue

        # 安全：開始需早於結束（sidebar 已擋，但這裡再保護一次）
        if s >= e:
            continue

        skip_rules.append({"user": user_str, "t_start": s, "t_end": e})

    return skip_rules


def _ensure_session_defaults():
    if "qc_last_result" not in st.session_state:
        st.session_state.qc_last_result = None
    if "qc_last_filename" not in st.session_state:
        st.session_state.qc_last_filename = None


def _split_am_pm(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    依「時段」欄位分割 AM/PM。支援：上午/下午
    """
    df = df.copy()
    if "時段" not in df.columns:
        return pd.DataFrame(), pd.DataFrame()

    df["班別"] = df["時段"].replace({"上午": "AM 班", "下午": "PM 班"})
    am_df = df[df["班別"] == "AM 班"].copy()
    pm_df = df[df["班別"] == "PM 班"].copy()
    return am_df, pm_df


def _safe_sum(df: pd.DataFrame, col: str) -> float:
    if df is None or df.empty or col not in df.columns:
        return 0.0
    return float(pd.to_numeric(df[col], errors="coerce").fillna(0).sum())


def _safe_mean(df: pd.DataFrame, col: str) -> float:
    if df is None or df.empty or col not in df.columns:
        return 0.0
    s = pd.to_numeric(df[col], errors="coerce")
    return float(s.mean())


def _safe_rate_ge(df: pd.DataFrame, col: str, target: float) -> float:
    if df is None or df.empty or col not in df.columns:
        return 0.0
    s = pd.to_numeric(df[col], errors="coerce")
    return float((s >= float(target)).mean())


def _render_shift_block(title: str, sdf: pd.DataFrame, *, top_n: int, target: float):
    # KPI 卡
    card_open(f"{title}｜KPI")
    if sdf is None or sdf.empty:
        st.info("本班別無資料")
        card_close()
        return

    render_kpis(
        [
            KPI("人數", f"{len(sdf):,}"),
            KPI("總驗收筆數", f"{_safe_sum(sdf, '筆數'):,.0f}"),
            KPI("總工時", f"{_safe_sum(sdf, '總工時'):.2f}"),
            KPI("平均效率", f"{_safe_mean(sdf, '效率'):.2f}"),
            KPI("達標率", f"{_safe_rate_ge(sdf, '效率', target):.0%}"),
        ],
        cols=5,
    )
    card_close()

    # 表格卡（整列紅/綠）
    card_open(f"{title}｜KPI 明細（紅=未達標）")
    # 這裡用 common_ui.show_kpi_table：效率 < target 整列紅
    show_kpi_table(sdf, eff_col="效率", target=float(target))
    card_close()

    # TopN 圖表卡
    card_open(f"{title}｜效率排行（Top {top_n}）")
    bar_topN(
        sdf,
        x_col="姓名" if "姓名" in sdf.columns else sdf.columns[0],
        y_col="效率",
        hover_cols=[c for c in ["筆數", "總工時", "空窗分鐘"] if c in sdf.columns],
        top_n=int(top_n),
        target=float(target),  # ✅ 低於 target 自動紅色
        title="低於門檻自動標紅；虛線為門檻",
    )
    card_close()


# =========================================================
# Main
# =========================================================
def main():
    _ensure_session_defaults()

    set_page(
        "驗收作業效能（KPI）",
        icon="✅",
        subtitle="驗收作業｜人時效率｜AM / PM 班別｜KPI 達標分析",
    )

    # ======================
    # Sidebar：計算條件設定（TopN + 排除空窗）
    # ======================
    controls = sidebar_controls(
        default_top_n=30,
        enable_exclude_windows=True,
        state_key_prefix="qc",
    )
    top_n = int(controls.get("top_n", 30))
    skip_rules = _adapt_exclude_windows_to_skip_rules(controls.get("exclude_windows", []))

    # ======================
    # 上傳資料 + 產出按鈕
    # ======================
    card_open("📤 上傳作業原始資料（驗收）", right_badge="XLSX / CSV")
    uploaded = st.file_uploader(
        "上傳驗收作業原始資料",
        type=["xlsx", "xls", "csv"],
        label_visibility="collapsed",
    )

    run_clicked = st.button("🚀 產出 KPI", type="primary", disabled=(uploaded is None))
    card_close()

    # ======================
    # 計算：點擊後寫入 session_state（避免 rerun 清畫面）
    # ======================
    if run_clicked:
        with st.spinner("KPI 計算中，請稍候..."):
            result = run_qc_efficiency(
                uploaded.getvalue(),
                uploaded.name,
                skip_rules,
            )
        st.session_state.qc_last_result = result
        st.session_state.qc_last_filename = uploaded.name

    # ======================
    # 一律從 session_state 取用結果
    # ======================
    result = st.session_state.qc_last_result

    if not result:
        st.info("請先上傳驗收作業原始資料並點選「🚀 產出 KPI」")
        return

    df = result.get("ampm_df", pd.DataFrame())
    target = float(result.get("target_eff", 20.0))

    if df is None or df.empty:
        st.warning("已計算完成，但沒有產出可顯示的資料（ampm_df 為空）")
        return

    if "時段" not in df.columns:
        st.error("資料缺少『時段』欄位，無法區分 AM / PM 班別")
        return

    am_df, pm_df = _split_am_pm(df)

    # ======================
    # 匯出（放在上面：一行=按鈕，且畫面不會消失）
    # ======================
    if result.get("xlsx_bytes"):
        download_excel_card(
            result["xlsx_bytes"],
            result.get("xlsx_name", "驗收作業KPI.xlsx"),
            label="⬇️ 匯出 KPI 報表（Excel）",
        )

    # ======================
    # AM / PM 兩欄一致呈現
    # ======================
    col_l, col_r = st.columns(2, gap="large")

    with col_l:
        _render_shift_block("🌓 AM 班（驗收）", am_df, top_n=top_n, target=target)

    with col_r:
        _render_shift_block("🌙 PM 班（驗收）", pm_df, top_n=top_n, target=target)


if __name__ == "__main__":
    main()
