"""
QPSS 脚本フォーマッタ - 管理者ダッシュボード
サイドバーで管理者認証を行い、利用分析データを可視化する。
"""

import datetime

import pandas as pd
import streamlit as st

from analytics import fetch_all_logs


def check_admin_auth() -> bool:
    """URLパラメータ ?admin=1 がある場合のみサイドバーにログイン欄を表示する。"""
    params = st.query_params
    if params.get("admin") != "1" and not st.session_state.get("admin_authenticated"):
        return False

    with st.sidebar:
        st.markdown("---")
        st.caption("管理者メニュー")

        if st.session_state.get("admin_authenticated"):
            st.success("ログイン中", icon="🔓")
            if st.button("ログアウト", key="admin_logout"):
                st.session_state.admin_authenticated = False
                st.rerun()
            return True

        password = st.text_input("パスワード", type="password", key="admin_pw")
        if st.button("ログイン", key="admin_login"):
            admin_pw = st.secrets.get("admin", {}).get("password", "")
            if admin_pw and password == admin_pw:
                st.session_state.admin_authenticated = True
                st.rerun()
            else:
                st.error("パスワードが違います")
        return False


def render_dashboard():
    """分析ダッシュボードを描画する。"""
    st.divider()
    st.header("利用分析ダッシュボード")

    logs = fetch_all_logs()
    if not logs:
        st.info("まだ分析データがありません。利用が記録されると、ここにグラフが表示されます。")
        return

    df = pd.DataFrame(logs)

    # 型変換
    df["timestamp"] = pd.to_datetime(df["timestamp"], errors="coerce")
    df["hour"] = pd.to_numeric(df["hour"], errors="coerce").fillna(0).astype(int)
    df["date"] = pd.to_datetime(df["date"], errors="coerce")

    # --- 期間フィルタ ---
    period = st.radio(
        "表示期間",
        ["直近7日", "直近30日", "全期間"],
        horizontal=True,
        key="dashboard_period",
    )
    today = pd.Timestamp(datetime.date.today())
    if period == "直近7日":
        df = df[df["date"] >= today - pd.Timedelta(days=6)]
    elif period == "直近30日":
        df = df[df["date"] >= today - pd.Timedelta(days=29)]

    if df.empty:
        st.info("選択期間内のデータがありません。")
        return

    # --- KPI カード ---
    total_views = len(df[df["event"] == "page_view"])
    total_converts = len(df[df["event"] == "convert"])
    today_views = len(
        df[(df["event"] == "page_view") & (df["date"] == today)]
    )
    convert_rate = (
        f"{total_converts / total_views * 100:.1f}%"
        if total_views > 0
        else "-"
    )

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("総アクセス数", total_views)
    col2.metric("総変換数", total_converts)
    col3.metric("本日のアクセス", today_views)
    col4.metric("変換率", convert_rate)

    # --- 日別アクセス推移 ---
    st.subheader("日別アクセス推移")
    daily = (
        df[df["event"].isin(["page_view", "convert"])]
        .groupby(["date", "event"])
        .size()
        .unstack(fill_value=0)
    )
    # カラム名を日本語に
    rename_map = {"page_view": "アクセス", "convert": "変換"}
    daily = daily.rename(columns=rename_map)
    st.line_chart(daily)

    # --- 時間帯別利用率 ---
    st.subheader("時間帯別利用分布")
    hourly = (
        df[df["event"] == "page_view"]
        .groupby("hour")
        .size()
        .reindex(range(24), fill_value=0)
    )
    hourly.index = [f"{h}時" for h in range(24)]
    st.bar_chart(hourly)

    # --- ファイル形式の内訳 ---
    uploads = df[df["event"] == "file_upload"]
    if not uploads.empty:
        st.subheader("ファイル形式の内訳")
        fmt_counts = uploads["file_format"].value_counts()
        col_chart, col_table = st.columns([2, 1])
        with col_chart:
            st.bar_chart(fmt_counts)
        with col_table:
            st.dataframe(
                fmt_counts.rename("件数").to_frame(),
                use_container_width=True,
            )

    # --- プリセット利用ランキング ---
    converts = df[df["event"] == "convert"]
    if not converts.empty:
        st.subheader("プリセット利用ランキング")
        preset_counts = converts["preset_name"].value_counts()
        st.dataframe(
            preset_counts.rename("使用回数").to_frame(),
            use_container_width=True,
        )

    # --- 生データ ---
    with st.expander("生データ"):
        st.dataframe(
            df.sort_values("timestamp", ascending=False),
            use_container_width=True,
        )
