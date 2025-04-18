# Kintone ユーザー・グループ管理ツール

Kintoneのユーザーとグループ情報を効率的に管理するためのPythonツールセットです。各スクリプトは特定の機能を担当し、ユーザーおよびグループデータの取得、整理、出力を行います。

## 主な機能

### 1. ユーザー・グループ情報のExcel出力

`get_user_group.py`を使用してKintoneの全ユーザーと各ユーザーの所属グループをExcelファイルに出力します。以下の機能を備えています：

- ✨ アクティブユーザーと停止中ユーザーを別シートで管理
- 👥 グループ所属状況を「●」で一目で確認
- ✉️ ログイン名とメールアドレスの整合性チェック
- 🔑 管理者（Administrators）は太字で強調表示

#### 主なクラスとメソッド

```python:get_user_group.py
class KintoneClient:
    def __init__(self, subdomain: str, username: str, password: str, logger: logging.Logger):
        ...
    
    def get_all_users(self) -> List[Dict[str, Any]]:
        ...
    
    def get_all_groups(self) -> List[Dict[str, Any]]:
        ...
    
    def get_users_in_group(self, group_code: str) -> List[Dict[str, Any]]:
        ...
```

- **KintoneClient**: Kintone APIとの通信を担当します。ユーザーやグループのデータを取得するメソッドを提供します。

```python:get_user_group.py
class DataProcessor:
    def __init__(self, users: List[Dict[str, Any]], groups: List[Dict[str, Any]], client: KintoneClient, logger: logging.Logger):
        ...
    
    def filter_groups(self) -> List[Dict[str, Any]]:
        ...
    
    def organize_groups(self, filtered_groups: List[Dict[str, Any]]) -> List[str]:
        ...
    
    def map_users(self):
        ...
    
    def populate_group_memberships(self, filtered_groups: List[Dict[str, Any]]):
        ...
    
    def generate_dataframes(self, group_names: List[str]) -> Dict[str, pd.DataFrame]:
        ...
```

- **DataProcessor**: 取得したユーザーおよびグループデータを処理し、Excel出力用のデータフレームを生成します。

```python:get_user_group.py
class ExcelExporter:
    def __init__(self, dataframes: Dict[str, pd.DataFrame], group_names: List[str], output_file: str, logger: logging.Logger):
        ...
    
    def export_to_excel(self):
        ...
    
    def format_excel(self):
        ...
```

- **ExcelExporter**: データフレームをExcelファイルに出力し、フォーマットを適用します。

### 2. グループ情報の管理

`get_group_list.py`でKintoneの全グループ情報をYAML形式で出力します。主な機能は以下の通りです：

- グループ名とIDの対応表を`group_list.yaml`として保存
- 簡単な検索・参照が可能

#### 主なクラスとメソッド

```python:get_group_list.py
class KintoneClient:
    def __init__(self, subdomain: str, username: str, password: str, logger: logging.Logger):
        ...
    
    def get_all_groups(self) -> List[Dict[str, Any]]:
        ...
```

- **KintoneClient**: ユーザー情報取得時と同様に、グループ情報の取得を担当します。

### 3. ユーザー情報の管理

`get_user_list.py`でKintoneの全ユーザー情報をYAML形式で出力します。主な機能は以下の通りです：

- 詳細なユーザー情報を`user_list.yaml`として保存
- データの再利用や他システムとの連携が容易

#### 主なクラスとメソッド

```python:get_user_list.py
class KintoneClient:
    def __init__(self, subdomain: str, username: str, password: str, logger: logging.Logger):
        ...
    
    def get_all_users(self) -> List[Dict[str, Any]]:
        ...
```

- **KintoneClient**: ユーザー情報の取得を担当します。

## セットアップと使用方法

### 必要なパッケージ

プロジェクトで必要なPythonパッケージは`requirements.txt`に記載されています。以下のコマンドでインストールしてください：

```bash
pip install -r requirements.txt
```

### 認証設定

ツールを使用するにはKintoneの認証情報が必要です。以下の2つの方法から選択できます：

1. **コマンドライン引数による設定**

   ```bash
   python get_user_group.py -d [ドメイン名] -u [ユーザー名] -p [パスワード]
   ```

2. **設定ファイルによる設定**

   `config_UserAccount.yaml`を作成し、以下の形式で認証情報を記述します：

   ```yaml:config_UserAccount.yaml
   subdomain: your-subdomain
   username: your-username
   password: your-password
   ```

   デフォルトでは、各スクリプトは`config_UserAccount.yaml`を参照します。

### 実行方法

各スクリプトの実行方法は以下の通りです：

```bash
# ユーザー・グループ情報をExcelに出力
python get_user_group.py

# グループ一覧をYAMLに出力
python get_group_list.py

# ユーザー一覧をYAMLに出力
python get_user_list.py
```

### 出力ファイル

- **`kintone_users_groups.xlsx`**: ユーザーとグループの情報が整理されたExcelファイル。
- **`group_list.yaml`**: グループ名とIDの対応表。
- **`user_list.yaml`**: 詳細なユーザー情報の一覧。

## ファイル構成

プロジェクトのディレクトリ構成は以下の通りです：

```
.
├── README.md
├── requirements.txt
├── config_UserAccount.yaml
├── get_user_group.py
├── get_group_list.py
└── get_user_list.py
```

## 詳細なプロセス

### 認証情報の取得

各スクリプトは以下の順序で認証情報を取得します：

1. コマンドライン引数から取得。
2. 引数が不足している場合、`config_UserAccount.yaml`から読み込み。
3. パスワードが未指定の場合、プロンプトで入力。

### データの取得

- **ユーザー情報**: `KintoneClient.get_all_users()`メソッドを使用して全ユーザー情報を取得。
- **グループ情報**: `KintoneClient.get_all_groups()`メソッドを使用して全グループ情報を取得。
- **グループ内ユーザー情報**: `KintoneClient.get_users_in_group(group_code)`メソッドを使用して各グループ内のユーザー情報を取得。

### データの処理

- **DataProcessor**クラスがユーザーとグループのデータを統合し、必要な情報を整理します。
- アクティブユーザーと停止中ユーザーを分け、所属グループを一覧化します。
- ログイン名とメールアドレスの整合性をチェックし、相違がある場合は「相違」列に表示します。

### Excelファイルへの出力とフォーマット

- **ExcelExporter**クラスが整理されたデータフレームをExcelファイルに出力します。
- ヘッダーのスタイル設定、列幅の調整、特定グループ（Administrators）のユーザー名を太字にするなどのフォーマットを適用します。

## 注意事項

- **パッケージのインストール**: 実行前に必要なPythonパッケージをインストールしてください。
- **認証情報の管理**: 認証情報は安全に管理し、`config_UserAccount.yaml`は`.gitignore`に追加することを推奨します。
- **ファイルの存在確認**: スクリプト実行時に必要なファイル（例：設定ファイル）が存在することを確認してください。

## ライセンス

MIT License

---

このツールセットを活用して、Kintoneのユーザーおよびグループ管理を効率化しましょう。問題や改善点があれば、ぜひフィードバックをお寄せください。
