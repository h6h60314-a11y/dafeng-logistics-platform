from __future__ import annotations

import hashlib
import json
import re
import uuid
from datetime import datetime
from typing import Any, Mapping, Optional, Sequence

import pandas as pd
import streamlit as st


UPLOAD_HEADERS = [
    "created_at",
    "page",
    "field",
    "file_name",
    "size_bytes",
    "sha256",
    "note",
]

RUN_HEADERS = [
    "run_id",
    "created_at",
    "page",
    "status",
    "parameters_json",
    "summary_json",
    "result_rows",
    "result_sheet",
    "download_filename",
]

EFFICIENCY_HEADERS = [
    "saved_at",
    "work_date",
    "department",
    "operation",
    "employee_id",
    "employee_name",
    "shift",
    "work_count",
    "work_minutes",
    "efficiency",
    "target",
    "is_pass",
    "source_page",
    "run_id",
    "note",
]

SYSTEM_SHEETS = {"uploads", "runs", "efficiency_daily", "Sheet1", "工作表1"}


def _now_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _safe_sheet_name(value: str, prefix: str = "result") -> str:
    text = re.sub(r"[\[\]\*\?/\\:]", "_", str(value or "data")).strip()
    text = re.sub(r"\s+", "_", text)
    text = text[:80].strip("_") or "data"
    name = f"{prefix}_{text}" if prefix else text
    return name[:100]


def _json(value: Any) -> str:
    return json.dumps(value or {}, ensure_ascii=False, default=str)


def sha256_bytes(content: bytes) -> str:
    return hashlib.sha256(content).hexdigest()


def is_configured() -> bool:
    return bool(st.secrets.get("GSHEET_ID")) and bool(_service_account_info())


def _service_account_info() -> Optional[dict[str, Any]]:
    if "google_service_account" in st.secrets:
        return dict(st.secrets["google_service_account"])

    raw = st.secrets.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if raw:
        return json.loads(raw)

    return None


def service_account_email() -> str:
    info = _service_account_info() or {}
    return str(info.get("client_email", ""))


@st.cache_resource(show_spinner=False)
def _spreadsheet():
    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except Exception as exc:
        raise RuntimeError(
            "尚未安裝 Google Sheet 套件，請在 requirements.txt 加入 gspread 與 google-auth。"
        ) from exc

    info = _service_account_info()
    sheet_id = st.secrets.get("GSHEET_ID")
    if not info or not sheet_id:
        raise RuntimeError("尚未設定 GSHEET_ID 或 google_service_account。")

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    credentials = Credentials.from_service_account_info(info, scopes=scopes)
    client = gspread.authorize(credentials)
    return client.open_by_key(sheet_id)


def _worksheet(title: str, headers: Sequence[str]):
    spreadsheet = _spreadsheet()
    try:
        ws = spreadsheet.worksheet(title)
    except Exception:
        ws = spreadsheet.add_worksheet(title=title, rows=1000, cols=max(20, len(headers)))
        ws.append_row(list(headers))
        return ws

    first_row = ws.row_values(1)
    if not first_row:
        ws.append_row(list(headers))
    elif list(first_row) != list(headers):
        merged = list(first_row)
        for header in headers:
            if header not in merged:
                merged.append(header)
        ws.update("1:1", [merged])
    return ws


def _clear_read_caches():
    for func in (_cached_sheet_records, _cached_worksheet_names):
        try:
            func.clear()
        except Exception:
            pass


def append_record(sheet_name: str, record: Mapping[str, Any], headers: Sequence[str]):
    ws = _worksheet(sheet_name, headers)
    current_headers = list(headers)
    row = [record.get(header, "") for header in current_headers]
    ws.append_row(row, value_input_option="USER_ENTERED")
    _clear_read_caches()


def ensure_database() -> bool:
    return _ensure_database_cached()


@st.cache_data(ttl=600, show_spinner=False)
def _ensure_database_cached() -> bool:
    try:
        _worksheet("uploads", UPLOAD_HEADERS)
        _worksheet("runs", RUN_HEADERS)
        _worksheet("efficiency_daily", EFFICIENCY_HEADERS)
    except PermissionError as exc:
        email = service_account_email()
        raise PermissionError(
            "Google Sheet 權限不足。請到 Google Sheet 按「共用」，"
            f"把 service account 加為「編輯者」：{email}"
        ) from exc
    return True


def append_efficiency_records(records: Sequence[Mapping[str, Any]], *, run_id: str = "") -> int:
    if not records:
        return 0

    rows = []
    saved_at = _now_text()
    for record in records:
        row = dict(record)
        row.setdefault("saved_at", saved_at)
        row.setdefault("run_id", run_id)
        rows.append(row)

    ws = _worksheet("efficiency_daily", EFFICIENCY_HEADERS)
    current_headers = list(EFFICIENCY_HEADERS)
    values = [[row.get(header, "") for header in current_headers] for row in rows]
    ws.append_rows(values, value_input_option="USER_ENTERED")
    _clear_read_caches()
    return len(rows)


def save_efficiency_dataframe(
    df: pd.DataFrame,
    *,
    source_page: str,
    department: str = "",
    operation: str = "",
    date_col: str = "日期",
    employee_id_col: str = "",
    employee_name_col: str = "姓名",
    shift_col: str = "",
    count_col: str = "筆數",
    minutes_col: str = "總分鐘",
    efficiency_col: str = "效率",
    target: Optional[float] = None,
    note: str = "",
) -> str:
    if df is None or df.empty:
        return ""

    run_id = uuid.uuid4().hex[:12]
    records: list[dict[str, Any]] = []
    for _, row in df.iterrows():
        if date_col and date_col in df.columns:
            work_date = pd.to_datetime(row.get(date_col), errors="coerce")
            work_date_text = work_date.strftime("%Y-%m-%d") if pd.notna(work_date) else str(row.get(date_col, ""))
        else:
            work_date_text = ""

        if not work_date_text:
            continue

        employee_id = str(row.get(employee_id_col, "")) if employee_id_col and employee_id_col in df.columns else ""
        employee_name = str(row.get(employee_name_col, "")) if employee_name_col in df.columns else employee_id
        efficiency = pd.to_numeric(pd.Series([row.get(efficiency_col, "")]), errors="coerce").iloc[0] if efficiency_col in df.columns else None
        work_count = row.get(count_col, "") if count_col in df.columns else ""
        work_minutes = row.get(minutes_col, "") if minutes_col in df.columns else ""
        shift = str(row.get(shift_col, "")) if shift_col and shift_col in df.columns else ""
        row_target = target
        if row_target is None and "達標門檻" in df.columns:
            row_target = pd.to_numeric(pd.Series([row.get("達標門檻", "")]), errors="coerce").iloc[0]
        target_value = float(row_target) if row_target is not None and pd.notna(row_target) else ""
        is_pass = ""
        if target_value != "" and pd.notna(efficiency):
            is_pass = bool(float(efficiency) >= float(target_value))

        records.append(
            {
                "work_date": work_date_text,
                "department": department,
                "operation": operation,
                "employee_id": employee_id,
                "employee_name": employee_name,
                "shift": shift,
                "work_count": work_count,
                "work_minutes": work_minutes,
                "efficiency": "" if pd.isna(efficiency) else float(efficiency),
                "target": target_value,
                "is_pass": is_pass,
                "source_page": source_page,
                "note": note,
            }
        )

    append_efficiency_records(records, run_id=run_id)
    return run_id


def record_upload(uploaded_file: Any, *, page: str, field: str = "", note: str = "") -> dict[str, Any]:
    if uploaded_file is None:
        return {}

    content = uploaded_file.getvalue()
    record = {
        "created_at": _now_text(),
        "page": page,
        "field": field,
        "file_name": getattr(uploaded_file, "name", ""),
        "size_bytes": len(content),
        "sha256": sha256_bytes(content),
        "note": note,
    }
    append_record("uploads", record, UPLOAD_HEADERS)
    return record


def record_upload_once(uploaded_file: Any, *, page: str, field: str = "", note: str = "") -> dict[str, Any]:
    if uploaded_file is None:
        return {}

    content = uploaded_file.getvalue()
    digest = sha256_bytes(content)
    state_key = f"gs_upload_saved::{page}::{field}::{getattr(uploaded_file, 'name', '')}::{digest}"
    if st.session_state.get(state_key):
        return {}

    record = {
        "created_at": _now_text(),
        "page": page,
        "field": field,
        "file_name": getattr(uploaded_file, "name", ""),
        "size_bytes": len(content),
        "sha256": digest,
        "note": note,
    }
    append_record("uploads", record, UPLOAD_HEADERS)
    st.session_state[state_key] = True
    return record


def record_uploads(files: Any, *, page: str, field: str = "", note: str = "") -> list[dict[str, Any]]:
    if files is None:
        return []
    if not isinstance(files, list):
        files = [files]
    return [record_upload(file, page=page, field=field, note=note) for file in files if file is not None]


def record_uploads_once(files: Any, *, page: str, field: str = "", note: str = "") -> list[dict[str, Any]]:
    if files is None:
        return []
    if not isinstance(files, list):
        files = [files]
    return [record_upload_once(file, page=page, field=field, note=note) for file in files if file is not None]


def save_run(
    *,
    page: str,
    status: str = "success",
    parameters: Optional[Mapping[str, Any]] = None,
    summary: Optional[Mapping[str, Any]] = None,
    result_rows: int = 0,
    result_sheet: str = "",
    download_filename: str = "",
    run_id: Optional[str] = None,
) -> str:
    run_id = run_id or uuid.uuid4().hex[:12]
    append_record(
        "runs",
        {
            "run_id": run_id,
            "created_at": _now_text(),
            "page": page,
            "status": status,
            "parameters_json": _json(parameters),
            "summary_json": _json(summary),
            "result_rows": result_rows,
            "result_sheet": result_sheet,
            "download_filename": download_filename,
        },
        RUN_HEADERS,
    )
    return run_id


def save_result_dataframe(
    *,
    page: str,
    df: pd.DataFrame,
    parameters: Optional[Mapping[str, Any]] = None,
    summary: Optional[Mapping[str, Any]] = None,
    download_filename: str = "",
    worksheet_name: Optional[str] = None,
    max_rows: int = 5000,
) -> str:
    if df is None:
        df = pd.DataFrame()

    run_id = uuid.uuid4().hex[:12]
    result_sheet = worksheet_name or _safe_sheet_name(page)
    limited = df.head(max_rows).copy()
    limited.insert(0, "_page", page)
    limited.insert(0, "_saved_at", _now_text())
    limited.insert(0, "_run_id", run_id)

    values_df = limited.fillna("").astype(str)
    headers = list(values_df.columns)
    ws = _worksheet(result_sheet, headers)
    current_headers = ws.row_values(1) or headers
    for header in headers:
        if header not in current_headers:
            current_headers.append(header)
    if current_headers != ws.row_values(1):
        ws.update("1:1", [current_headers])

    rows = []
    for _, row in values_df.iterrows():
        rows.append([row.get(header, "") for header in current_headers])
    if rows:
        ws.append_rows(rows, value_input_option="USER_ENTERED")
        _clear_read_caches()

    save_run(
        run_id=run_id,
        page=page,
        status="success",
        parameters=parameters,
        summary=summary,
        result_rows=len(df),
        result_sheet=result_sheet,
        download_filename=download_filename,
    )
    return run_id


def read_sheet(sheet_name: str, max_rows: int = 500) -> pd.DataFrame:
    try:
        rows = _cached_sheet_records(sheet_name, int(max_rows))
    except Exception:
        ensure_database()
        rows = _cached_sheet_records(sheet_name, int(max_rows))
    if not rows:
        return pd.DataFrame()
    return pd.DataFrame(rows)


@st.cache_data(ttl=60, show_spinner=False)
def _cached_sheet_records(sheet_name: str, max_rows: int) -> list[dict[str, Any]]:
    ws = _spreadsheet().worksheet(sheet_name)
    rows = ws.get_all_records()
    if not rows:
        return []
    return rows[-int(max_rows):]


def list_worksheets() -> list[str]:
    ensure_database()
    return _cached_worksheet_names()


@st.cache_data(ttl=120, show_spinner=False)
def _cached_worksheet_names() -> list[str]:
    return [ws.title for ws in _spreadsheet().worksheets()]


def list_result_worksheets() -> list[str]:
    return [name for name in list_worksheets() if name not in SYSTEM_SHEETS and name.startswith("result_")]


def read_efficiency_daily(start_date: Any, end_date: Any, max_rows: int = 5000) -> pd.DataFrame:
    df = read_sheet("efficiency_daily", max_rows=max_rows)
    if df is None or df.empty or "work_date" not in df.columns:
        return pd.DataFrame()

    data = df.copy()
    data["_work_date"] = pd.to_datetime(data["work_date"], errors="coerce")
    start = pd.to_datetime(start_date)
    end = pd.to_datetime(end_date)
    data = data[(data["_work_date"] >= start) & (data["_work_date"] <= end)]
    data = data.drop(columns=["_work_date"])
    return data
