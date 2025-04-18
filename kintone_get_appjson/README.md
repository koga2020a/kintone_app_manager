# kintone アプリ情報取得ツール

## 前提条件

0. 必要なモジュール
pip install requests
pip install openpyxl 
pip install pyyaml
pip install bs4



1. 以下のファイルを作成してください：

### config_UserAccount.yaml
```yaml
subdomain: <your-subdomain>
username: <your-username>
password: <your-password>
```

### run_scripts_params.tsv
```tsv
アプリidの番号  APIキー
```
※タブ区切りで、アプリIDとAPIキーを記載してください。

### group_list.yaml / user_list.yaml
以下のコマンドで生成できます：
```bash
# グループ一覧の取得
python ..\kintone_get_user_group\get_group_list.py --config ..\kintone_get_user_group\config_UserAccount.yaml

# ユーザー一覧の取得
python ..\kintone_get_user_group\get_user_list.py --config ..\kintone_get_user_group\config_UserAccount.yaml
```

## 使用方法

### 1. アプリ情報のダウンロード

```bash
# 全アプリの情報を取得
python run_scripts.py -a

# 特定のアプリ情報を取得（run_scripts_params.tsv内の番号を指定）
python run_scripts.py <アプリ番号>
```

または直接実行する場合：
```bash
python download2yaml_excel.py <アプリ番号> <APIキー>
```

出力先：`.\<アプリ番号>_【アプリ名】_日時\`

### 2. ACL情報のExcel変換

```bash
# 基本的な使用方法
python aclJson_to_excel.py <アプリ番号>

# ログレベルを指定して実行
python aclJson_to_excel.py <アプリ番号> --log-level {DEBUG|INFO|WARNING|ERROR|CRITICAL}

# サイレントモードで実行（ログ出力を抑制）
python aclJson_to_excel.py <アプリ番号> --silent

# グループマスタファイルのパスを指定
python aclJson_to_excel.py <アプリ番号> --group-master <パス>
```

オプション:
- `--log-level`: ログの出力レベルを指定（デフォルト: INFO）
  - DEBUG: 詳細なデバッグ情報を出力
  - INFO: 一般的な情報を出力
  - WARNING: 警告メッセージを出力
  - ERROR: エラーメッセージを出力
  - CRITICAL: 重大なエラーメッセージのみ出力
- `--silent`: すべてのログ出力を抑制
- `--group-master`: グループマスタファイルのパスを指定（デフォルト: ../kintone_get_user_group/group_user_list.yaml）

### 出力ファイル

#### aclJson_to_excel.py の出力
- アクセス権限情報をExcel形式に変換（`<アプリ番号>_acl.xlsx`）
  - レコードシート：レコードレベルのアクセス権限
  - アプリシート：アプリレベルのアクセス権限
- ユーザー名一覧（`<アプリ番号>permission_target_user_names.csv`）
  - アクセス権限に関連する全ユーザーの一覧

### 3. ACL問題レポートの生成

```bash
# 基本的な使用方法
python make_all_acl_problem_report.py

# 出力ディレクトリとファイル名を指定して実行
python make_all_acl_problem_report.py -d ./output -o custom_report.csv

# 詳細なログ出力を有効にして実行
python make_all_acl_problem_report.py -v
```

オプション:
- `-d`, `--directory`: 探索する基準ディレクトリを指定（デフォルト: ./output）
- `-o`, `--output`: 出力ファイル名を指定（デフォルト: all_acl_problem_report.csv）
- `-v`, `--verbose`: 詳細なログ出力を有効にする

出力ファイル（タブ区切りCSV）の内容:
- ヘッダ名: アプリID
- メッセージ: 問題の内容または「問題未検出」
- タイプ: ユーザー/グループの区分
- 名称: ユーザー名またはグループ名
- 矛盾タイプ: 権限の矛盾内容
- 出現回数: 問題の発生回数
- 詳細: 追加情報
- ディレクトリ名: 対象アプリのディレクトリ

注意事項:
- 数字で始まるディレクトリ名のみが処理対象となります
- 問題が検出されなかったアプリは「問題未検出」として出力されます
- 結果は「問題あり」が上部に、「問題未検出」が下部にソートされて出力されます

## 注意事項

1. APIキーは以下の権限が必要です：
   - レコード閲覧
   - アプリ管理

2. 機密情報を含むファイルは.gitignoreに登録されています：
   - config_UserAccount.yaml
   - group_list.yaml
   - user_list.yaml
   - run_scripts_params.tsv

## その他の機能

### レイアウト情報の詳細出力
以下の手順で実行できます：

1. アプリ情報のダウンロード
```bash
python test_get_js_359.py <アプリ番号> <APIキー> <出力先ディレクトリ>
```

2. レイアウトデータの生成
```bash
python make_52_layout.py <ディレクトリパス>
```
出力：out01.tsv

3. データの下処理
```bash
python depLine_out_to_out2.py
```
出力：out02.tsv

4. Excelファイルの生成と罫線設定
```bash
python find_L_shaped_cells.py <Excelファイル名>
python draw_full_borders.py <Excelファイル名>
```

■グループcodeとグループ名の対照表としてlistを取得する
　⇒出力ファイル名： ./group_list.yaml    （実行ディレクトリに出力されます）
python ..\kintone_get_user_group\get_group_list.py --config ..\kintone_get_user_group\config_UserAccount.yaml

■ユーザnoとユーザ情報の対照表としてlistを取得する
　⇒出力ファイル名： ./group_list.yaml    （実行ディレクトリに出力されます）
python ..\kintone_get_user_group\get_user_list.py --config ..\kintone_get_user_group\config_UserAccount.yaml
s
■run_scripts.py
  tsvファイル:run_scripts_params.tsv のAPIキー情報で、download2yaml_excel.pyを実行します。
  使い方：
    全行：
      python download2yaml_excel.py -a
    アプリ番号指定：※run_scripts_params.tsv内の番号を指定します。
      python download2yaml_excel.py 288

■download2yaml_excel.py
  kintoneの指定アプリの情報を、アプリ番号のディレクトリに出力します。
  使い方：
    python download2yaml_excel.py アプリ番号 APIキー

  アプリ番号 288 のとき、出力先は .\288_【アプリ名】_日時 以下となります。

■指定のアプリidのディレクトリで、aclのyamlからエクセルでの表を生成します。
  python .\aclJson_to_excel.py 52

# YAML to Excel ACL Converter

## 概要

**YAML to Excel ACL Converter** は、YAML形式で記述されたアクセス制御リスト（ACL）ファイルをExcelファイルに変換するPythonスクリプトです。このスクリプトは、ユーザーやグループの権限情報を視覚的に管理しやすいExcel形式に整形します。特に、エンティティの種類が `GROUP` または `USER` の場合に、それぞれ「グループ」および「ユーザ」として表示されるように設計されています。

## 特徴

- **YAMLからExcelへの変換**: ACL情報をわかりやすいExcelシートに変換。
- **エンティティタイプの日本語表示**: `USER` タイプのエンティティは「ユーザ」と表示。
- **カスタマイズ可能なマスタファイル**: グループやユーザーのマスタ情報を外部ファイルから読み込み。
- **条件フィルタのサポート**: レコードの条件に基づいた権限設定を反映。
- **エラーチェックと警告**: 無効なグループやユーザーに対する警告表示。

## 必要条件

- **Python 3.7 以降**
- 以下のPythonライブラリ
  - `PyYAML`
  - `openpyxl`

## インストール

1. **Pythonのインストール**

   Pythonがインストールされていない場合は、[公式サイト](https://www.python.org/downloads/)からインストールしてください。

2. **依存ライブラリのインストール**

   以下のコマンドを実行して、必要なPythonライブラリをインストールします。

   ```bash
   pip install PyYAML openpyxl

step1_DownloadAppJson.py    appid、app token、出力先ディレクトリ
  APPのjson、yaml多数
  必須：__52__form_layout.json、__52__form_fields.json

step2_MakeLayoutFromJson.py
  out01.tsv

step3_DropLine_1to2.py
  out02.tsv

step4_SeiriLine_2toExcel.py
  out03_excel_test.xlsx

APIでkintoneの情報を取得します。

loop_app_check.py
    概要：kintone rest api で取得できる情報を取得します。（全体対象）
        ※jsのダウンロードに関してのみ、レコード閲覧、アプリ管理が両方ともtrueのものだけが対象です。
    Usage: python loop_app_check.py <base_directory>
    例: python loop_app_check.py ../
    数字のみのディレクトリ内の _s__apitoken.yaml から、レコード閲覧・アプリ管理の両方のAPI権限があるものについて、処理を実行します。

python test_get_js_359.py 
    概要：kintone rest api で取得できる情報を取得します。（個別対象）
    Usage: python script.py <appid> <api_token> <base_directory>
    例: python script.py 288 NO7M0FxwFiqgagxHVcCaZqJ3VGrjd9krT6v2oqe8 ../

◆画面情報をエクセルに生成
過程が複数あります：
１．アプリの情報をkintoneサイトからダウンロードする。
　　python test_get_js_359.py にてyaml、json取得（アプリのapiキーでアプリ編集許可のキーを利用するもの）


２．１番から取得した情報（ファイル）から、レイアウト用のデータを生成する。
　　　　　例：画面に表示される準に、フィールド名・フィールドコード、必須項目設定の有無などが羅列
　　python make_52_layout.py [ディレクトリpath:52]
　　　試験用の制限：52番の名前のファイルのみ対応
　　　　　　出力は実行ディレクトリに out01.tsv
⇒これから新規にexcelを生成する。
⇒レイアウトは既存のと同等にして、スクリプトで生成する。

２-2．２番のout01.tsvを、下処理してout02.tsvにする　（そのまま転記可能なレベルにする）
　　python depLine_out_to_out2.py
　　　試験用の制限：out01.tsv
　　　　　　出力：out02.tsv


３．【手動】excel.xlsx に、out01.tsvを転記する
⇒書き込み先を仕様にすれば、次への連携が可能

４．labelやsubtableのグループ番号をセルに書き込む（L字になる）
　　python find_L_shaped_cells.py [エクセルファイル名]
　　　試験用の制限：excel.xlsx
　　　　　　出力：配列の配列の文字列
⇒このまま５番に連携可能

５．４番でL字の配列から、L字形状の罫線を引く（文字色、背景色の設定も可能）
　　python draw_full_borders.py [エクセルファイル名]
　　　試験用の制限：excel.xlsx
　　　　　　出力：excel2.xlsx



　　　　　　　　　　
