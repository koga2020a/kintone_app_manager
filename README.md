# kintone アプリ管理ツール

kintone のアプリ設定・フォーム・アクセス権 (ACL)・通知・プロセス管理・カスタマイズ JS を
まとめてダウンロードし、YAML / JSON で保存したうえで Excel の一覧表に変換するツールです。
取得済みデータの全文検索や、ユーザー・グループの一覧出力にも対応しています。

## 基本の流れ

1. **設定** — `.kintone.env` に接続先とアプリごとの API トークンを書く
2. **ダウンロード** — アプリ設定と JS を `output/[アプリID]_[アプリ名]/` に取得する
3. **Excel 変換** — 取得済みデータから ACL・通知・プロセス管理・設定一覧表を作る
4. **検索・確認** — 取得済みの YAML / JSON / JS を全文検索して内容を確かめる

対話メニュー (`python menu.py`) なら 1 → 2 → 3 → 4 を番号を選ぶだけで実行できます。

ACL・通知の Excel は、`users` が作るマスタ YAML (`output/group_user_list.yaml` /
`output/user_list.yaml`) を使ってグループ名・ユーザー名を解決します。名前を反映させるため、
**`users` → `app` → `acl` / `notifications`** の順に実行してください（`all` はこの順で実行します）。

## ドメイン変更の as-is / to-be 確認

サーバのドメイン名を変える前後で、Webhook の通知先・JavaScript 内の HTTP 呼び出し先・
外部参照 URL に古いホストが残っていないかを、取得済みの全アプリから横断して確認できます。

1. **対応表を登録** — `.kintone.env` の `domain_map` に `変更前ホスト: 変更後ホスト` を書く
   （対話メニューの「設定 → ドメイン対応表」からも登録できます）

   ```yaml
   domain_map:
     old.example.com: "new.example.com"
   ```

2. **最新を取得** — `python kintone_runner.py app` で対象アプリを取得し直す
3. **一覧を出す** — `python kintone_runner.py urls` で `output/url_inventory_[日時].xlsx` を作り、
   **要変更** の行を確認して関係者と共有する
4. **変更後に再確認** — ドメイン変更が終わったら再度 `app` → `urls` を実行し、
   **要変更が 0 件**で、対象の URL が **変更済** になっていることを確認する

`domain_map` が未設定でも一覧は出せます（「変更後 URL」「状態」の列が空になります）。

### Excel のシート構成

| シート | 内容 |
| --- | --- |
| サマリ | ホストごとの集計（Webhook / JS呼び出し / 外部参照 の件数、出現アプリ、状態、変更後ホスト） |
| Webhook | 各アプリの Webhook 通知先 URL |
| JS呼び出し | JavaScript 内の HTTP 呼び出し（`fetch` / `XMLHttpRequest` / `kintone.proxy` / `$.ajax` など）の URL |
| 外部参照 | カスタマイズ設定で URL 指定した JS / CSS と、カスタムビューの HTML 内にある URL |

### 状態

| 状態 | 意味 | やること |
| --- | --- | --- |
| 要変更 | ホストが `domain_map` の**変更前ホスト**と一致する | ドメイン変更に合わせて直す |
| 変更済 | ホストが `domain_map` の**変更後ホスト**と一致する | 対応済み |
| 対象外 | `domain_map` に無いホスト（外部サービスなど） | 変更不要 |

> **JS の検出は best-effort です。** JavaScript 内の URL は静的解析（ファイルを実行せずに読む方式）で
> 拾っているため、文字列リテラルで書かれた URL は検出できますが、変数経由で渡している URL や
> テンプレートリテラル・文字列連結で組み立てている URL は `(動的)` と表示されたり、式のままの形で
> 出ることがあります。0 件でも「無い」とは限らないので、気になる箇所は `search` でも確認してください。

検出の主な制限:

- 走査するのは `javascript/` 配下の `*.js` だけです。1 行が 2000 文字を超えるファイル（ミニファイ済み）は
  ライブラリとみなして丸ごとスキップし、サマリの「スキップしたファイル数」に数えます。
- 変数の解決は同じファイル内の `const/let/var` 定義 1 段階だけです。他ファイルの定義や関数の戻り値は追いません。
- `kintone.api(...)` は kintone 自身への呼び出しなのでホストを `(kintone)` とし、状態は常に「対象外」です。
- `domain_map` の照合はホスト名だけで、ポートやパスは見ません。`old.example.com` を登録すると
  `api.old.example.com` のようなサブドメインも「要変更」になり、その部分だけ置き換えた変更後 URL を出します。
- Webhook は取得済みファイルから読みます。`app` を username / password 付きで実行していないアプリは
  「Webhook 未取得」として警告し、集計に含まれません。
- 通知の宛先、プラグイン設定、アクションの URL、レコード内のデータは対象外です。

## 必要なもの

- Python 3.8 以上
- 依存パッケージ

```bash
pip install requests pyyaml pandas openpyxl
```

仮想環境を使う場合は `python -m venv venv` で作成し、Windows は `. venv/Scripts/activate`、
macOS / Linux は `source venv/bin/activate` で有効化してから上記の `pip install` を実行してください。

## 設定ファイル `.kintone.env`

スクリプトと同じフォルダに YAML 形式で置きます。

```yaml
# 同じフォルダの .kintone.env に、次の形式で書いてください
subdomain: "your_subdomain"    # example.cybozu.com の example 部分
username: "admin@example.com"  # 管理者のログイン名
password: "your_password"
user_domain: "example.com"     # 任意。ユーザー一覧のメインドメイン

# ここが対象アプリの登録箇所（アプリID: APIトークン）
app_tokens:
  123: "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
  456: "yyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyy"

# 任意。ドメイン変更の as-is / to-be 確認用（変更前ホスト: 変更後ホスト）
domain_map:
  old.example.com: "new.example.com"
```

| 項目 | 必須 | 内容 |
| --- | --- | --- |
| `subdomain` | 必須 | `example.cybozu.com` の `example` 部分 |
| `username` | 必須 | 管理者のログイン名 |
| `password` | 必須 | 上記ユーザーのパスワード |
| `app_tokens` | 必須 | 対象アプリの `アプリID: APIトークン` |
| `user_domain` | 任意 | ユーザー一覧のメインドメイン |
| `search_keywords` | 任意 | `search` で検索語を省略したときに使う既定の検索語（リスト） |
| `domain_map` | 任意 | ドメイン変更の対応表（`変更前ホスト: 変更後ホスト`）。`urls` の「変更後 URL」「状態」に使う |
| `js_dirs` | 任意 | ローカルの JS フォルダ（`名前: パス`）。フィールドコードの使用箇所を追加で調べる |

設定は対話メニューの「12. 設定」からも編集できます。

### API トークンの用意

1. kintone で対象アプリを開く
2. 「アプリの設定」→「設定」タブ →「APIトークン」
3. 生成し、**[必須] レコード閲覧** / **[推奨] アプリ管理** の権限を付ける
4. 画面右上の「**アプリを更新**」で反映する（これを忘れると 403 になります）
5. `.kintone.env` の `app_tokens` に `<アプリID>: "<発行したトークン>"` を追加

### 認証について

kintone の認証は、API トークンを `X-Cybozu-API-Token` ヘッダで送る方法と、
ユーザー名・パスワードによる Basic 認証の 2 種類です。
このツールでは、アプリ系 (`app` / `acl` / `notifications` / `process_workflow` / `summary`) は
API トークンを、`users` / `group` / `webhooks` はユーザー名・パスワードを使います。

## 使い方 A: 対話メニュー

```bash
python menu.py
```

```
======================================================================
  kintone アプリ管理ツール
======================================================================
  接続先     https://your_subdomain.cybozu.com/  (admin@example.com)
  対象アプリ 1 件: 4 顧客リスト (取得 08/19 10:05)
  設定ファイル .kintone.env  … 変更は 12

  取得
   1. アプリ設定・JS をダウンロード
  Excel 変換（1 で取得したデータから作成）
   2. アクセス権 (ACL)
   3. 通知設定
   4. プロセス管理
   5. アプリ設定一覧表（全アプリ）
  閲覧・検索
   6. 取得済みデータを検索（JS 含む）
   7. Webhook 一覧
   8. URL 一覧を Excel 出力（Webhook / JS 呼び出し / 外部参照）
  ユーザー・グループ
   9. ユーザー一覧を Excel 出力
  10. グループ操作（一覧・検索・追加・削除）
  その他
  11. まとめて実行（ユーザー一覧 → ダウンロード → Excel 変換）
  12. 設定（接続先 / アプリとトークン / 検索文字列）
  13. ヘルプ（使い方・設定ファイル・出力ファイル）
   0. 終了

番号を選択 [0-13]:
```

`python menu.py --detail` を使うと各項目の説明と出力先も表示し、内部の `kintone_runner.py` を `--verbose` で実行します。

## 使い方 B: CLI

```bash
python kintone_runner.py <コマンド> [オプション]
```

| コマンド | 内容 | 主なオプション |
| --- | --- | --- |
| `users` | ユーザーと所属グループを取得 | `--format {excel,csv}` |
| `app` | アプリ設定・フォーム・ACL・通知・JS をダウンロード | `--id` |
| `acl` | アクセス権 (ACL) を Excel に変換 | `--id` |
| `summary` | 全アプリの設定一覧表を出力 | `--output` |
| `notifications` | 通知設定を Excel に変換 | `--id` |
| `process_workflow` | プロセス管理を Excel に変換 | `--id` |
| `search` | 取得済みの YAML / JSON / JS を全文検索 | `[検索語...]` `--id` `--no-excel` |
| `webhooks` | Webhook 設定を一覧表示（未取得なら API で取得） | `--id`（必須） |
| `urls` | Webhook・JS の HTTP 呼び出し・外部参照 URL を一覧化して Excel 出力 | `--id`（複数可） `--output` |
| `group` | グループ操作 `list` / `search` / `add` / `remove` | — |
| `all` | `users` → `app` → `acl` → `summary` → `notifications` → `process_workflow` | `--id` `--not-id` |
| `outputs` | 生成されるファイルの一覧と概要を表示 | — |

`--id` を省略したコマンドは `app_tokens` に登録した全アプリが対象になります。

共通オプション（サブコマンドの前・後どちらに置いても動作します）

- `--env PATH` : 使用する `.kintone.env` のパス
- `-v`, `--verbose` : INFO ログも端末に表示する（`search` は内訳も表示）

### 例

```bash
# 生成されるファイルの一覧を確認する（API は呼ばない）
python kintone_runner.py outputs

# 登録済みの全アプリの設定・JS をダウンロード
python kintone_runner.py app

# アプリID 123 だけをダウンロード（前回分は previous_output/ に退避）
python kintone_runner.py app --id 123

# 取得済みデータから ACL の Excel を作成
python kintone_runner.py acl --id 123

# 取得済みデータを全文検索（Excel も出力。検索語を省略すると search_keywords を使用）
python kintone_runner.py search kintone.events getFieldValue --id 123

# 検索して端末に表示するだけ（Excel を作らない）
python kintone_runner.py search 顧客コード --no-excel

# Webhook の設定を一覧表示
python kintone_runner.py webhooks --id 123

# 取得済み全アプリの URL（Webhook / JS 呼び出し / 外部参照）を一覧化
python kintone_runner.py urls

# アプリを絞って一覧化（複数のアプリIDを並べて指定できます）
python kintone_runner.py urls --id 123 456

# ユーザー一覧を CSV で出力
python kintone_runner.py users --format csv

# グループ操作
python kintone_runner.py group list
python kintone_runner.py group search "山田"
python kintone_runner.py group add "user_code" "グループ名"

# 別の設定ファイルを使ってまとめて実行（アプリID 123 を除外）
python kintone_runner.py --env /path/to/other.env all --not-id 123
```

### 終了コード

| コード | 意味 |
| --- | --- |
| 0 | 成功 |
| 1 | 一部の処理が失敗 |
| 2 | 引数エラー |
| 3 | API トークンの権限不足（レコード閲覧を付けて「アプリを更新」が必要） |

## 出力ファイル

出力先は `./output/` です。アプリ単位のファイルは `output/[アプリID]_[アプリ名]/` の中に入ります。
（`python kintone_runner.py outputs` で同じ一覧を確認できます）

| ファイル | 内容 | 生成コマンド |
| --- | --- | --- |
| `kintone_users_groups_[日時].xlsx` | ユーザーとグループの一覧 | `users`（既定） |
| `kintone_users_groups_[日時].csv` | 同上（CSV・UTF-8 BOM 付き） | `users --format csv` |
| `[アプリID]_acl_report.xlsx` | アプリ・レコード・フィールドのアクセス権（ユーザー名・グループ名を反映） | `acl` |
| `[アプリID]permission_target_user_names.csv` | アプリに出現するユニークなユーザー名一覧 | `acl` |
| `kintone_app_settings_summary_[日時].xlsx` | 全アプリの設定一覧表 | `summary` |
| `[アプリID]_notifications.xlsx` | 通知設定（一般・レコード・リマインダー） | `notifications` |
| `[アプリID]_process_workflow.xlsx` | プロセス管理（ステータスと作業者） | `process_workflow` |
| `[アプリID]_layout_report.xlsx` | フォームレイアウトの一覧表 | `app`（取得時に自動生成） |
| `[アプリID]_layout_raw.tsv` / `_layout_structured.tsv` | フォームレイアウトの生データ / 整形データ | `app`（取得時に自動生成） |
| `search_hits_[日時].xlsx` | 全文検索の結果（検索語・アプリ毎のサマリとヒット一覧） | `search`（`--no-excel` で抑止） |
| `url_inventory_[日時].xlsx` | Webhook・JS の HTTP 呼び出し・外部参照 URL の横断一覧（`domain_map` があれば変更後 URL と状態も） | `urls` |

`users` / `app` / `acl` / `summary` / `notifications` / `process_workflow` は `all` でまとめて生成できます。

### ディレクトリ

| パス | 内容 |
| --- | --- |
| `audit/` | kintone の監査ログ CSV / ZIP を置く入力フォルダ。`users` の実行時に無ければ作成され、置いてあれば最終ログイン日の集計に使われる |
| `output/` | 最新の取得結果。`output/[アプリID]_[アプリ名]/` に YAML、`json/`、`javascript/`、Excel が入る。`users` が作るマスタ YAML (`group_user_list.yaml` など) もここ |
| `previous_output/` | `all` および `app --id` の実行時に、前回の取得結果を退避する場所 |
| `backup/[日時]/` | 実行後に `output/` をコピーしたもの |
| `logs/` | 実行ログ `kintone_runner_[日時].log` |
| `error_report.txt` | 失敗したときのエラー詳細（追記されていく） |

## ファイル構成

| ファイル / ディレクトリ | 役割 |
| --- | --- |
| `menu.py` | 対話メニュー。主な入口。内部で `kintone_runner.py` を呼び出す |
| `kintone_runner.py` | CLI 本体。各サブコマンドの実行と、ログ・バックアップ・終了コードの管理 |
| `console.py` | 端末表示の共通ヘルパー（色・記号・見出し） |
| `inspect_downloaded.py` | 取得済みデータの全文検索と、検索結果の Excel 出力 |
| `url_inventory.py` | Webhook・JS 内の URL 抽出と、URL 一覧 Excel (`urls`) の出力 |
| `app_settings_summary.py` | アプリ設定一覧表 (`summary`) の生成 |
| `.kintone.env` | 設定ファイル（YAML） |
| `kintone_get_appjson/` | アプリ設定・JS のダウンロード本体 |
| `kintone_get_user_group/` | ユーザー・グループ情報の取得本体 |
| `kintone_group_cli/` | グループ操作の本体 |
| `tests/` | unittest のテスト（`python -m unittest discover -s tests`） |
| `kintone_get_appjson/make_all_acl_problem_report.py` | 各アプリの `*_acl_problem.csv` を横断集計する単体ツール（runner からは呼ばれない） |

`kintone_get_appjson` / `kintone_get_user_group` / `kintone_group_cli` は
`kintone_runner.py` から子プロセスとして呼び出される実体スクリプトです。

## トラブルシューティング

- **終了コード 3 / 403 が出る** — API トークンの権限不足です。対象アプリの「APIトークン」で
  レコード閲覧（必要に応じてアプリ管理）を付け、**「アプリを更新」を押して反映**してから再実行してください。
- **Excel の書き込みでエラーになる** — 対象の `.xlsx` を Excel で開いたままだと書き込めません。
  ファイルを閉じてから再実行してください。
- **詳細を知りたい** — `logs/kintone_runner_[日時].log` に全ログが残ります。
  失敗時の詳細は `error_report.txt` にも追記されます。端末に詳しく出したいときは `-v` を付けてください。

## 注意事項

- ユーザー・グループ関連の機能には kintone の**管理者権限**が必要です。
- `.kintone.env` にはパスワードと API トークンが入ります。**Git にコミットしないでください**
  （`.gitignore` で除外済みです）。
- 対象アプリが多い場合や `all` の実行には時間がかかることがあります。
