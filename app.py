from __future__ import annotations

import ast
from pathlib import Path

import streamlit as st

from app_pages import HIDDEN_NAV_LABELS, HIDDEN_NAV_URL_KEYS, NAV_SECTIONS, PageSpec
from common_ui import inject_sidebar_nav_style


st.set_page_config(
    page_title="大豐物流 - 作業平台",
    page_icon="assets/gf_logo.png",
    layout="wide",
)

inject_sidebar_nav_style(
    hidden_labels=HIDDEN_NAV_LABELS,
    hidden_url_keys=HIDDEN_NAV_URL_KEYS,
)

BROKEN_PAGES: list[tuple[str, str]] = []
MISSING_PAGES: list[str] = []


def _syntax_ok(path: Path) -> bool:
    try:
        ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        return True
    except Exception as exc:
        BROKEN_PAGES.append((str(path), repr(exc)))
        return False


def _page_if_available(spec: PageSpec):
    path = Path(spec.path)
    if not path.exists():
        MISSING_PAGES.append(spec.path)
        return None
    if not _syntax_ok(path):
        return None
    try:
        return st.Page(
            spec.path,
            title=spec.title,
            icon=spec.icon,
            url_path=spec.url_path,
            default=spec.default,
        )
    except Exception as exc:
        BROKEN_PAGES.append((spec.path, f"st.Page 建立失敗：{repr(exc)}"))
        return None


def _build_navigation():
    navigation = {}
    for section_name, specs in NAV_SECTIONS.items():
        pages = [_page_if_available(spec) for spec in specs]
        navigation[section_name] = [page for page in pages if page]
    return navigation


def _show_preflight_warnings():
    if MISSING_PAGES:
        with st.sidebar.expander("⚠️ 找不到檔案（未載入）", expanded=False):
            st.caption("下列 pages 檔案不存在，所以不會出現在左側選單：")
            for path in MISSING_PAGES:
                st.code(path)

    if BROKEN_PAGES:
        with st.sidebar.expander("⚠️ 已停用頁面（語法/縮排錯）", expanded=True):
            st.caption("以下檔案有 IndentationError / SyntaxError，已自動略過避免整站掛掉：")
            for path, error in BROKEN_PAGES:
                st.code(f"{path}\n{error}")


navigation = _build_navigation()
_show_preflight_warnings()

pg = st.navigation(navigation, expanded=False)
pg.run()
