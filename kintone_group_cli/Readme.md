# kintone Group CLI

kintone のグループ管理を行うためのコマンドラインツールです。

## 機能

- ユーザーの検索
- グループ一覧の表示
- ユーザーのグループ設定（追加/削除）

## 必要要件

- Python 3.7以上
- pip

## インストール

1. 必要なパッケージをインストールします：

```bash
pip install requests pandas openpyxl pyyaml
```

2. 設定ファイル `.kintone.env` を作成します：

```yaml
subdomain: "your-subdomain"
username: "your-username"
password: "your-password"
app_tokens:
  "1": "token-for-app-1"
  "2": "token-for-app-2"
```

## 使い方

### ユーザー検索

ログイン名、表示名、メールアドレスに対して部分一致で検索を行います：

```bash
python group_cli.py 検索キーワード
```

検索結果が表示され、以下の動作をします：
- 検索結果が1名の場合：自動的にそのユーザーの所属グループ一覧を表示
- 検索結果が複数名の場合：番号を選択することでそのユーザーの所属グループ一覧を表示
  
表示される情報：
- ユーザーの基本情報（ログイン名、表示名、メールアドレス）
- 選択したユーザーの所属グループ一覧（システムグループを除く）

例：
```bash
$ python group_cli.py john

検索結果:
--------------------------------------------------------------------------------
No.  ログイン名                      表示名                          メールアドレス
--------------------------------------------------------------------------------
1    johndoe                        John Doe                        john@example.com
2    johnson                        Johnson Smith                   johnson@example.com
--------------------------------------------------------------------------------
合計: 2件

詳細を表示するユーザーの番号を入力してください (0でキャンセル): 1

John Doe (johndoe) の所属グループ:
--------------------------------------------------
1. Developers (コード: dev_team)
2. Project A (コード: project_a)
--------------------------------------------------

### グループ一覧の表示

```bash
python group_cli.py list
```

### ユーザーのグループ設定

1. グループを直接指定する場合：

```bash
python group_cli.py set ユーザー名 グループ名
```

2. 対話的にグループを選択する場合：

```bash
python group_cli.py set ユーザー名
```

※ グループ設定時は、指定したグループ以外のグループからユーザーが削除されます。

### オプション

- `--config`: 設定ファイルのパスを指定（デフォルト: config_UserAccount.yaml）
- `--silent`: 詳細なログを表示しない
- `--search`: ユーザー検索モードを有効にする（第一引数を検索キーワードとして扱う）

```bash
python group_cli.py --config path/to/config.yaml list
python group_cli.py --search set  # "set"をキーワードとしてユーザーを検索
```

## 設定ファイル

`.kintone.env` の形式：

```yaml
subdomain: "your-subdomain"    # kintoneのサブドメイン（example.kintone.com の example 部分）
username: "your-username"      # 管理者権限を持つユーザーのログイン名
password: "your-password"      # パスワード
app_tokens:                    # アプリIDとAPIトークンのマッピング
  "1": "token-for-app-1"
  "2": "token-for-app-2"
```

## 注意事項

- 管理者権限を持つユーザーの認証情報が必要です
- グループの設定を変更する場合は、十分に注意して実行してください
- パスワードなどの認証情報は、安全に管理してください
- 動的グループのグループコードは指定できません
- Excelファイルが開かれている状態での実行はエラーとなります

## エラー対応

1. 認証エラー
   - 設定ファイルの認証情報が正しいか確認してください
   - ユーザーが必要な権限を持っているか確認してください

2. ネットワークエラー
   - インターネット接続を確認してください
   - kintone が利用可能な状態か確認してください

3. 設定ファイルの読み込みエラー
   - 設定ファイルのパスが正しいか、ファイルが存在するか確認してください
   - YAML形式が正しいか確認してください

4. Excelファイルの操作エラー
   - 出力先のExcelファイルが開かれていないか確認してください
   - 十分なディスク容量があるか確認してください

## ライセンス

このプロジェクトはMITライセンスの下でライセンスされています。詳細はLICENSEファイルを参照してください。

## ファイル構成

### `group_cli.py`

```python:group_cli.py
import argparse
import sys
import base64
from getpass import getpass
import logging
import yaml
import requests

class KintoneClient:
    # クラスの詳細

class GroupManager:
    # クラスの詳細

def setup_logging(silent: bool = False, debug: bool = False) -> logging.Logger:
    # ロギング設定

def load_config(config_path: str = 'config_UserAccount.yaml') -> dict:
    # 設定ファイルの読み込み

def main():
    # エントリーポイント

if __name__ == "__main__":
    main()
```

## 使用例

### ユーザーを検索する

```bash
python group_cli.py john
python group_cli.py --search list  # "list"をキーワードとして検索
```

### グループ一覧を表示する

```bash
python group_cli.py list
```

### ユーザーを特定のグループに設定する

```bash
python group_cli.py set johndoe Developers
```

### ユーザーのグループを対話的に設定する

```bash
python group_cli.py set johndoe
```

### デバッグログを有効にしてグループ一覧を表示する

```bash
python group_cli.py --debug list
```