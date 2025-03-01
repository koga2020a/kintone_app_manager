# Kintone統合実行ツール

このツールは、Kintone関連の複数のユーティリティを1つのインターフェースから利用できるようにする統合スクリプトです。

## 機能

このツールは以下の機能を提供します：

1. **ユーザーとグループ情報の取得** (`kintone_get_user_group`)  
   Kintoneの全ユーザーと各ユーザーの所属グループ情報をExcelファイルに出力します。

2. **アプリのJSONデータ取得** (`kintone_get_appjson`)  
   指定したアプリのフィールド設定やACL情報などをJSON形式で取得します。

3. **グループ操作** (`kintone_group_cli`)  
   グループの一覧取得、ユーザー検索、グループへのユーザー追加/削除を行います。

## 必要条件

- Python 3.6以上
- 必要なPythonパッケージ：
  - requests
  - pyyaml
  - pandas
  - openpyxl

## インストール

1. リポジトリをクローンするか、ファイルをダウンロードします。

2. 必要なパッケージをインストールします：

```bash
pip install requests pyyaml pandas openpyxl
```

3. 設定ファイルを準備します。`.kintone.env`ファイルを以下の形式で作成します：

```yaml
# Kintoneのサブドメイン (example.kintone.com の example 部分)
subdomain: "your_subdomain" 

# 管理者権限を持つユーザーのログイン名
username: "admin@example.com" 

# パスワード
password: "your_password"

# アプリIDとAPIトークンのマッピング (オプション)
app_tokens:
  # app_id: "api_token"
  # 例:
  # 123: "abcdefg123456789hijklmn" 
```

## 使い方

### 基本的な使い方

```bash
python kintone_runner.py [コマンド] [オプション]
```

利用可能なコマンド：

- `users`: ユーザーとグループ情報を取得
- `app`: アプリのJSONデータを取得
- `group`: グループ操作
- `all`: すべての機能を順番に実行

### 具体的な使用例

#### ユーザーとグループ情報の取得

```bash
# デフォルト形式（Excel）で出力
python kintone_runner.py users

# 形式を指定（excel または csv）
python kintone_runner.py users --format excel
```

#### アプリのJSONデータ取得

```bash
# 設定ファイルに登録されている全アプリを処理
python kintone_runner.py app

# 特定のアプリIDを指定
python kintone_runner.py app --id 123
```

#### グループ操作

```bash
# グループ一覧を表示
python kintone_runner.py group list

# ユーザー検索
python kintone_runner.py group search "検索キーワード"

# ユーザーをグループに追加
python kintone_runner.py group add "ユーザーコード" "グループ名"

# ユーザーをグループから削除
python kintone_runner.py group remove "ユーザーコード"
```

#### すべての機能を実行

```bash
python kintone_runner.py all
```

#### 別の設定ファイルを使用

```bash
python kintone_runner.py [コマンド] --env /path/to/config.env
```

## 出力ファイル

- ユーザーとグループ情報：`output/kintone_users_groups_YYYYMMDD_HHMMSS.xlsx`
- アプリのJSONデータ：各アプリのIDなどに基づいた名前で`output/`ディレクトリに出力されます
- ログファイル：`logs/kintone_runner_YYYYMMDD_HHMMSS.log`

## 連携の仕組み

このツールは、以下のディレクトリの機能を連携して実行します：

1. `kintone_get_user_group`: ユーザーとグループ情報を取得して出力します
2. `kintone_get_appjson`: アプリのフィールド設定やACL情報を取得します
3. `kintone_group_cli`: グループ操作を行います

## トラブルシューティング

問題が発生した場合は、`logs/`ディレクトリ内のログファイルを確認してください。詳細なエラー情報が記録されています。

## 注意事項

- このツールを使用するには、Kintoneの管理者権限が必要です。
- APIトークンを使用する機能は、適切な権限を持つトークンを設定ファイルに指定する必要があります。
- 大量のデータを処理する場合は、処理に時間がかかることがあります。 