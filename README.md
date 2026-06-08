# QPSS 脚本フォーマッタ

脚本・シナリオのテキストを、規定の体裁（縦書き・横向き、MS Gothic 12pt、ト書きの自動インデントなど）に整えた **Word文書（.docx）** に変換するツールです。

メモ帳・Word・各種エディタで書いた原稿をアップロードするだけで、提出用フォーマットに揃った docx をダウンロードできます。ブラウザだけで使える Web 版と、コマンドライン版の両方に対応しています。

---

## 主な機能

- **複数の入力形式に対応**: `.txt` / `.docx` / `.rtf` / `.md`（Markdown 装飾は自動で除去）
- **ト書きの自動判定**: セリフ・柱・テロップ・ナレーション・登場人物紹介などを除外し、残りをト書きとみなして TAB インデントを付与
- **YAMLプリセット**: ページの向き・文字方向・用紙サイズ・フォント・インデント・行間・ページ番号・執筆者表記などをまとめて切り替え
- **マイプリセット**: 詳細設定をブラウザに保存／YAML でエクスポート・インポート（Web版）
- **執筆者名の挿入**: 「脚本：〇〇」をタイトル直後・ヘッダー・非表示から選択
- **利用分析ダッシュボード**: 管理者向けに利用状況を可視化（任意設定）

---

## 1. Web版（最も簡単）

ブラウザで以下にアクセスして使います。インストール不要です。

👉 **https://qpss-scenario-formatter-fdrkrqerpcxwyd3wxoslmq.streamlit.app/**

### 手順

1. **脚本ファイルをアップロード**（`.txt` / `.docx` / `.rtf` / `.md`）
2. 必要なら **執筆者名** と **出力ファイル名** を入力
3. **プリセットを選択**（例: 「QPSSフォーマッタ標準」＝縦書き・横向き）
4. （任意）「詳細設定」を開いて、向き・フォント・インデント・行間などを微調整
5. **「Word文書を生成」** をクリック
6. **「ダウンロード」** で docx を保存

> 💡 アップロードしたファイルはサーバーのメモリ上でのみ処理され、保存されません。詳しくは [SECURITY.md](SECURITY.md) を参照してください。機密性の高い脚本は、後述の **ローカル実行** を推奨します。

---

## 2. ローカルで実行する

インターネットを経由せず、自分のPC内だけで処理したい場合はこちらを使います。

### セットアップ

```bash
git clone https://github.com/nanana-monokaki/qpss-scenario-formatter.git
cd qpss-scenario-formatter

# 仮想環境（任意・推奨）
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux: source .venv/bin/activate

pip install -r requirements.txt
```

### Web UI をローカルで起動

```bash
streamlit run app.py
```

ブラウザが自動で開きます（通常 http://localhost:8501 ）。操作方法は Web版と同じです。

---

## 3. コマンドライン（CLI）で使う

GUIを使わず、ターミナルから一括変換したい場合は `format_scenario.py` を直接実行します。

```bash
# プリセット一覧を表示
python format_scenario.py --list-presets

# 変換（標準プリセット）
python format_scenario.py -i 原稿.txt -o 完成稿.docx -w "山田太郎"

# 横書きプリセットを使う
python format_scenario.py --preset horizontal -i 原稿.txt
```

### オプション

| オプション | 説明 | 既定値 |
|-----------|------|--------|
| `--preset` | プリセット名（拡張子なし）またはYAMLファイルのパス | `default` |
| `-i`, `--input` | 入力ファイルパス | （必須） |
| `-o`, `--output` | 出力ファイルパス | `<入力名>_formatted.docx` |
| `-w`, `--writer` | 執筆者名（「脚本：〇〇」として挿入） | （なし） |
| `--list-presets` | 利用可能なプリセット一覧を表示 | — |

---

## プリセットについて

プリセットは `presets/` 配下の YAML ファイルです。

| ファイル | 内容 |
|---------|------|
| `default.yaml` | QPSS標準（縦書き・横向き・A4） |
| `horizontal.yaml` | 横書き・縦向きの一般的な脚本フォーマット |
| `_template.yaml` | 新規プリセット作成用のテンプレート（`_` 始まりは一覧に出ません） |

### 新しいプリセットを追加する

`_template.yaml` をコピーして `presets/` に置けば、Web版・CLIの選択肢に自動で表示されます。主な設定項目は次のとおりです。

```yaml
name: "プリセット表示名"
description: "説明"

page:
  orientation: landscape       # landscape（横向き） / portrait（縦向き）
  text_direction: vertical     # vertical（縦書き） / horizontal（横書き）
  size: A4                     # A4 / B5 / Letter

font:
  name: "MS Gothic"
  size: 12                     # pt

paragraph:
  indent_before_mm: 6          # 前インデント(mm)
  indent_after_mm: 0
  space_before_pt: 0           # 段落前間隔(pt)
  space_after_pt: 0

header_footer:
  page_number: true
  page_number_position: right_bottom  # right_bottom / center_bottom / right_top
  page_number_font_size: 10

togaki:                        # ト書き判定
  indent_tabs: 2               # ト書きに付ける TAB 数
  auto_detect: true
  exclude_patterns:            # 「ト書きではない」行のパターン（下記参照）
    - { type: contains, value: "「" }

line_spacing:
  between_dialogue: 0              # セリフ⇔セリフ間の空行数
  between_togaki_and_dialogue: 1   # ト書き⇔セリフ間の空行数

meta:
  writer_label: "脚本"            # 「脚本：〇〇」の「脚本」部分
  writer_position: after_title    # after_title / header / none
```

### ト書き判定の仕組み

`exclude_patterns` に該当**しない**行が「ト書き」とみなされ、TABインデントが付きます。除外パターンの `type` は次の5種類です。

| type | 意味 |
|------|------|
| `contains` | `value` を含む行（例: `「` を含む＝セリフ） |
| `starts_with` | `value` で始まる行（例: `Ｔ「`＝テロップ） |
| `exact` | `value` に完全一致する行（例: `＜了＞`） |
| `regex` | 正規表現に一致する行（例: `^[０-９]`＝柱） |
| `char_name` | 「名前（年齢）」形式の登場人物紹介行（`max_paren_pos` で `（` の許容位置を指定） |

> 案件固有のルール（特定のタイトル行やロゴ表示など）は、各プリセットの `exclude_patterns` に追記して調整してください。

---

## 管理者ダッシュボード（任意）

利用状況の分析ダッシュボードを使う場合は、Streamlit の `secrets` を設定します（Google Sheets に匿名のイベントログを記録）。

`.streamlit/secrets.toml`（ローカル）または Streamlit Cloud の Secrets に以下を設定します。

```toml
[admin]
password = "管理者パスワード"

[analytics]
spreadsheet_id = "対象スプレッドシートのID"

[gcp_service_account]
# Google サービスアカウントのJSONキーの内容
type = "service_account"
project_id = "..."
private_key = "..."
client_email = "..."
# ...（以下、JSONキーの各項目）
```

設定後、アプリURLに `?admin=1` を付けてアクセスするとサイドバーにログイン欄が表示され、パスワード認証でダッシュボードを開けます。`secrets` 未設定でも分析機能が無効になるだけで、変換機能は通常どおり動作します。

---

## 必要環境

- Python 3.9 以上を推奨
- 依存パッケージは [requirements.txt](requirements.txt) を参照（`streamlit` / `python-docx` / `pyyaml` / `mammoth` / `striprtf` ほか）

---

## セキュリティ

アップロードファイルの扱いや運用上の注意点は [SECURITY.md](SECURITY.md) にまとめています。機密性の高い脚本はローカル実行を推奨します。
