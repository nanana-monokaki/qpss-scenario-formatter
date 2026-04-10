"""
QPSS 脚本フォーマッタ - 利用分析モジュール
Google Sheetsにイベントを記録し、管理者ダッシュボード用にデータを取得する。
"""

import datetime
import uuid

import streamlit as st

SHEET_NAME = "analytics_log"
COLUMNS = [
    "timestamp", "date", "hour", "event",
    "file_format", "preset_name", "preset_kind", "session_id",
]


def _get_session_id() -> str:
    """セッション固有の匿名IDを取得・生成する。"""
    if "analytics_session_id" not in st.session_state:
        st.session_state.analytics_session_id = str(uuid.uuid4())[:8]
    return st.session_state.analytics_session_id


@st.cache_resource(ttl=300)
def _get_worksheet():
    """Google Sheetsのワークシートをキャッシュして返す。
    認証情報が未設定または接続失敗時はNoneを返す。"""
    try:
        import gspread
        from google.oauth2.service_account import Credentials

        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
        ]
        creds_dict = dict(st.secrets.get("gcp_service_account", {}))
        if not creds_dict:
            return None

        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        gc = gspread.authorize(creds)

        spreadsheet_id = st.secrets.get("analytics", {}).get("spreadsheet_id", "")
        if not spreadsheet_id:
            return None

        sh = gc.open_by_key(spreadsheet_id)

        # シートが存在しなければ作成
        try:
            ws = sh.worksheet(SHEET_NAME)
        except gspread.exceptions.WorksheetNotFound:
            ws = sh.add_worksheet(title=SHEET_NAME, rows=1000, cols=len(COLUMNS))
            ws.append_row(COLUMNS, value_input_option="USER_ENTERED")

        return ws
    except Exception:
        return None


def log_event(
    event: str,
    file_format: str = "",
    preset_name: str = "",
    preset_kind: str = "",
):
    """イベントを1行追記する。失敗してもユーザー操作をブロックしない。"""
    try:
        now = datetime.datetime.utcnow()
        row = [
            now.isoformat(),
            now.strftime("%Y-%m-%d"),
            now.hour,
            event,
            file_format,
            preset_name,
            preset_kind,
            _get_session_id(),
        ]
        ws = _get_worksheet()
        if ws:
            ws.append_row(row, value_input_option="USER_ENTERED")
    except Exception:
        pass


def fetch_all_logs() -> list[dict]:
    """全ログデータを辞書のリストで返す。取得失敗時は空リストを返す。"""
    try:
        ws = _get_worksheet()
        if ws is None:
            return []
        records = ws.get_all_records()
        return records
    except Exception:
        return []
