"""
QPSS 脚本フォーマッタ - Web UI (Streamlit)
非エンジニア向けのシンプルなブラウザ操作画面。
"""

import json
import os
import sys
from pathlib import Path

import streamlit as st
import yaml
from streamlit_local_storage import LocalStorage

# スクリプトのあるディレクトリを基準にパスを解決
SCRIPT_DIR = Path(__file__).parent
sys.path.insert(0, str(SCRIPT_DIR))

from format_scenario import (
    ScenarioFormatter,
    extract_text,
    list_presets,
    load_preset,
    save_preset,
    PRESETS_DIR,
)
from analytics import log_event
from admin_dashboard import check_admin_auth, render_dashboard

# --- ブラウザ localStorage ---
local_storage = LocalStorage()
_STORAGE_KEY = "qpss_user_presets"


def _load_user_presets():
    """ブラウザ localStorage からユーザープリセットを読み込む"""
    raw = local_storage.getItem(_STORAGE_KEY)
    if raw and isinstance(raw, dict):
        return raw
    if raw and isinstance(raw, str):
        try:
            return json.loads(raw)
        except (json.JSONDecodeError, TypeError):
            return {}
    return {}


def _save_user_presets(presets_dict):
    """ユーザープリセットを localStorage に保存する"""
    local_storage.setItem(_STORAGE_KEY, presets_dict)

# --- ページ設定 ---
st.set_page_config(
    page_title="QPSS 脚本フォーマッタ",
    page_icon="📄",
    layout="centered",
)

st.title("📄 QPSS 脚本フォーマッタ")
st.caption("脚本ファイルをアップロードして、フォーマット済みのWord文書を生成します")

# ページビュー記録（1セッションにつき1回）
if "page_view_logged" not in st.session_state:
    log_event("page_view")
    st.session_state.page_view_logged = True

st.divider()

# --- プリセット読み込み ---
builtin_presets = list_presets()
user_presets = _load_user_presets()

if not builtin_presets and not user_presets:
    st.error("プリセットが見つかりません。presets/ フォルダにYAMLファイルを配置してください。")
    st.stop()

# 統合プリセットリスト: (表示名, 種別, キー)
preset_options = []
for p in builtin_presets:
    preset_options.append({
        "label": f"{p['name']} - {p['description']}",
        "kind": "builtin",
        "key": p["file"],
        "name": p["name"],
    })
for name in user_presets:
    preset_options.append({
        "label": f"{name}（マイプリセット）",
        "kind": "user",
        "key": name,
        "name": name,
    })

# --- メイン入力エリア ---
st.subheader("脚本ファイルをアップロード")
uploaded_file = st.file_uploader(
    "対応形式: .txt, .docx, .rtf, .md",
    type=["txt", "docx", "rtf", "md"],
    help="メモ帳、Word、その他エディタで作成した脚本ファイルをドラッグ＆ドロップまたは選択してください",
)

col1, col2 = st.columns(2)
with col1:
    writer_name = st.text_input("執筆者名", placeholder="例: 山田太郎")
with col2:
    output_filename = st.text_input(
        "出力ファイル名",
        value="脚本_formatted.docx",
        help="ダウンロードするファイルの名前",
    )

st.divider()

# --- プリセット選択 ---
st.subheader("フォーマット設定")

if not preset_options:
    st.error("利用可能なプリセットがありません。")
    st.stop()

selected_idx = st.selectbox(
    "プリセットを選択",
    range(len(preset_options)),
    format_func=lambda i: preset_options[i]["label"],
)

# 選択中のプリセットを読み込み
selected_option = preset_options[selected_idx]
if selected_option["kind"] == "builtin":
    current_preset = load_preset(selected_option["key"].replace(".yaml", ""))
else:
    current_preset = user_presets[selected_option["key"]].copy()

# --- 詳細設定（折りたたみ） ---
with st.expander("詳細設定（クリックで展開）"):
    st.caption("プリセットの設定を個別に変更できます。変更はこの実行のみ有効です。")

    tab1, tab2, tab_ls, tab3 = st.tabs(["ページ・フォント", "ト書き判定", "行間設定", "メタ情報"])

    with tab1:
        col_a, col_b = st.columns(2)
        with col_a:
            orientation = st.radio(
                "ページの向き",
                ["landscape", "portrait"],
                index=0 if current_preset.get("page", {}).get("orientation") == "landscape" else 1,
                format_func=lambda x: "横向き" if x == "landscape" else "縦向き",
                horizontal=True,
            )
            text_direction = st.radio(
                "文字方向",
                ["vertical", "horizontal"],
                index=0 if current_preset.get("page", {}).get("text_direction") == "vertical" else 1,
                format_func=lambda x: "縦書き" if x == "vertical" else "横書き",
                horizontal=True,
            )
            page_size = st.selectbox(
                "用紙サイズ",
                ["A4", "B5", "Letter"],
                index=["A4", "B5", "Letter"].index(
                    current_preset.get("page", {}).get("size", "A4")
                ),
            )

        with col_b:
            font_name = st.text_input(
                "フォント名",
                value=current_preset.get("font", {}).get("name", "MS Gothic"),
            )
            font_size = st.number_input(
                "フォントサイズ (pt)",
                value=current_preset.get("font", {}).get("size", 12),
                min_value=6,
                max_value=72,
            )

        col_c, col_d = st.columns(2)
        with col_c:
            indent_before = st.number_input(
                "前インデント (mm)",
                value=current_preset.get("paragraph", {}).get("indent_before_mm", 6),
                min_value=0,
                max_value=50,
            )
            space_before = st.number_input(
                "段落前間隔 (pt)",
                value=current_preset.get("paragraph", {}).get("space_before_pt", 0),
                min_value=0,
                max_value=72,
            )
        with col_d:
            indent_after = st.number_input(
                "後インデント (mm)",
                value=current_preset.get("paragraph", {}).get("indent_after_mm", 0),
                min_value=0,
                max_value=50,
            )
            space_after = st.number_input(
                "段落後間隔 (pt)",
                value=current_preset.get("paragraph", {}).get("space_after_pt", 0),
                min_value=0,
                max_value=72,
            )

        page_number = st.checkbox(
            "ページ番号を表示",
            value=current_preset.get("header_footer", {}).get("page_number", True),
        )
        if page_number:
            pn_position = st.selectbox(
                "ページ番号の位置",
                ["right_bottom", "center_bottom", "right_top"],
                index=["right_bottom", "center_bottom", "right_top"].index(
                    current_preset.get("header_footer", {}).get("page_number_position", "right_bottom")
                ),
                format_func=lambda x: {"right_bottom": "右下", "center_bottom": "中央下", "right_top": "右上"}[x],
            )
            pn_font_size = st.number_input(
                "ページ番号フォントサイズ (pt)",
                value=current_preset.get("header_footer", {}).get("page_number_font_size", 10),
                min_value=6,
                max_value=24,
            )
        else:
            pn_position = "right_bottom"
            pn_font_size = 10

        # 詳細設定の値をプリセットに反映
        current_preset["page"] = {
            "orientation": orientation,
            "text_direction": text_direction,
            "size": page_size,
        }
        current_preset["font"] = {"name": font_name, "size": font_size}
        current_preset["paragraph"] = {
            "indent_before_mm": indent_before,
            "indent_after_mm": indent_after,
            "space_before_pt": space_before,
            "space_after_pt": space_after,
        }
        current_preset["header_footer"] = {
            "page_number": page_number,
            "page_number_position": pn_position,
            "page_number_font_size": pn_font_size,
        }

    with tab2:
        st.caption("ト書きと判定されない行のパターンを設定します。")

        togaki_config = current_preset.get("togaki", {})
        auto_detect = st.checkbox("ト書き自動判定を有効にする", value=togaki_config.get("auto_detect", True))
        indent_tabs = st.number_input(
            "ト書きTABインデント数",
            value=togaki_config.get("indent_tabs", 2),
            min_value=0,
            max_value=8,
        )

        st.markdown("**除外パターン一覧**")
        patterns = togaki_config.get("exclude_patterns", [])

        # 既存パターンの表示・削除
        patterns_to_keep = []
        for i, pat in enumerate(patterns):
            pat_type = pat.get("type", "")
            pat_value = pat.get("value", "")
            display = f"{pat_type}: {pat_value}" if pat_value else f"{pat_type}"
            if pat_type == "char_name":
                display = f"char_name (max_paren_pos: {pat.get('max_paren_pos', 6)})"
            col_p, col_del = st.columns([5, 1])
            with col_p:
                st.text(display)
            with col_del:
                if st.button("削除", key=f"del_pat_{i}"):
                    continue
            patterns_to_keep.append(pat)

        # 新規パターン追加
        st.markdown("**パターンを追加**")
        col_type, col_val, col_add = st.columns([2, 3, 1])
        with col_type:
            new_type = st.selectbox(
                "タイプ",
                ["contains", "starts_with", "exact", "regex", "char_name"],
                label_visibility="collapsed",
            )
        with col_val:
            new_value = st.text_input("値", label_visibility="collapsed", placeholder="パターンの値を入力")
        with col_add:
            if st.button("追加"):
                if new_type == "char_name":
                    patterns_to_keep.append({"type": "char_name", "max_paren_pos": 6})
                elif new_value:
                    patterns_to_keep.append({"type": new_type, "value": new_value})

        current_preset["togaki"] = {
            "indent_tabs": indent_tabs,
            "auto_detect": auto_detect,
            "exclude_patterns": patterns_to_keep,
        }

    with tab_ls:
        st.caption("行種の切り替わり時に挿入する空行数を設定します。")
        ls_config = current_preset.get("line_spacing", {})
        col_ls1, col_ls2 = st.columns(2)
        with col_ls1:
            blank_between_dialogue = st.number_input(
                "セリフ間の空行数",
                value=ls_config.get("between_dialogue", 0),
                min_value=0,
                max_value=5,
                help="セリフとセリフの間に挿入する空行の数",
            )
        with col_ls2:
            blank_between_togaki = st.number_input(
                "ト書き⇔セリフ間の空行数",
                value=ls_config.get("between_togaki_and_dialogue", 1),
                min_value=0,
                max_value=5,
                help="ト書きとセリフの切り替わりに挿入する空行の数",
            )
        current_preset["line_spacing"] = {
            "between_dialogue": blank_between_dialogue,
            "between_togaki_and_dialogue": blank_between_togaki,
        }

    with tab3:
        meta_config = current_preset.get("meta", {})
        writer_label = st.text_input(
            "執筆者ラベル",
            value=meta_config.get("writer_label", "脚本"),
            help='「脚本：〇〇」の「脚本」部分を変更できます（例: 作、原案）',
        )
        writer_position = st.selectbox(
            "執筆者名の表示位置",
            ["after_title", "header", "none"],
            index=["after_title", "header", "none"].index(
                meta_config.get("writer_position", "after_title")
            ),
            format_func=lambda x: {"after_title": "タイトルの次の行", "header": "ヘッダー", "none": "表示しない"}[x],
        )
        current_preset["meta"] = {
            "writer_label": writer_label,
            "writer_position": writer_position,
        }

    # --- マイプリセット保存 ---
    st.divider()
    st.markdown("**現在の設定をマイプリセットとして保存**")
    st.caption("ブラウザに保存されます。他のユーザーには表示されません。")
    col_name, col_save = st.columns([3, 1])
    with col_name:
        save_name = st.text_input(
            "プリセット名",
            placeholder="例: 案件A用",
            label_visibility="collapsed",
        )
    with col_save:
        if st.button("保存", type="secondary"):
            if save_name:
                current_preset["name"] = save_name
                current_preset["description"] = f"マイプリセット: {save_name}"
                up = _load_user_presets()
                up[save_name] = current_preset
                _save_user_presets(up)
                st.success(f"マイプリセット「{save_name}」を保存しました")
                st.rerun()
            else:
                st.warning("プリセット名を入力してください")

    # --- マイプリセット削除 ---
    if selected_option["kind"] == "user":
        if st.button("このマイプリセットを削除", type="secondary"):
            up = _load_user_presets()
            up.pop(selected_option["key"], None)
            _save_user_presets(up)
            st.success(f"マイプリセット「{selected_option['name']}」を削除しました")
            st.rerun()

    # --- マイプリセット エクスポート / インポート ---
    st.divider()
    st.markdown("**マイプリセットのバックアップ**")
    st.caption("キャッシュクリアや別端末への移行に備えて、ファイルで保存・復元できます。")

    col_export, col_import = st.columns(2)

    with col_export:
        up_current = _load_user_presets()
        if up_current:
            export_yaml = yaml.dump(up_current, allow_unicode=True, default_flow_style=False)
            st.download_button(
                label="エクスポート（YAML）",
                data=export_yaml.encode("utf-8"),
                file_name="my_presets.yaml",
                mime="application/x-yaml",
                use_container_width=True,
            )
        else:
            st.info("保存済みのマイプリセットがありません")

    with col_import:
        import_file = st.file_uploader(
            "インポート（YAML）",
            type=["yaml", "yml"],
            key="preset_import",
            label_visibility="collapsed",
        )
        if import_file is not None:
            try:
                imported = yaml.safe_load(import_file.read().decode("utf-8"))
                if not isinstance(imported, dict):
                    st.error("無効なファイル形式です")
                else:
                    up_merged = _load_user_presets()
                    count = 0
                    for name, preset in imported.items():
                        if isinstance(preset, dict):
                            up_merged[name] = preset
                            count += 1
                    _save_user_presets(up_merged)
                    st.success(f"{count} 件のマイプリセットをインポートしました")
                    st.rerun()
            except Exception as e:
                st.error(f"インポートに失敗しました: {e}")

st.divider()

# --- 変換・ダウンロード ---
if uploaded_file is not None:
    # ファイルアップロード記録
    if "last_uploaded_file" not in st.session_state or st.session_state.last_uploaded_file != uploaded_file.name:
        log_event("file_upload", file_format=Path(uploaded_file.name).suffix.lower())
        st.session_state.last_uploaded_file = uploaded_file.name

    # テキスト抽出
    try:
        file_bytes = uploaded_file.read()
        text = extract_text(file_bytes, filename=uploaded_file.name)
    except Exception as e:
        st.error(f"ファイルの読み込みに失敗しました: {e}")
        st.stop()

    # プレビュー
    with st.expander("テキストプレビュー（抽出結果）", expanded=False):
        preview_lines = text.split("\n")[:50]
        st.text("\n".join(preview_lines))
        if len(text.split("\n")) > 50:
            st.caption(f"... 他 {len(text.split(chr(10))) - 50} 行")

    # 生成ボタン
    if st.button("Word文書を生成", type="primary", use_container_width=True):
        with st.spinner("Word文書を生成中..."):
            try:
                formatter = ScenarioFormatter(current_preset)
                docx_bytes = formatter.create_docx(text, writer_name=writer_name)

                log_event(
                    "convert",
                    preset_name=selected_option["name"],
                    preset_kind=selected_option["kind"],
                )
                st.success("Word文書の生成が完了しました")
                st.download_button(
                    label="ダウンロード",
                    data=docx_bytes,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"Word文書の生成に失敗しました: {e}")
                st.exception(e)
else:
    st.info("脚本ファイルをアップロードしてください")

# --- 管理者ダッシュボード ---
if check_admin_auth():
    render_dashboard()
