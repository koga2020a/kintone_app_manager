import csv
import subprocess
import sys
import os
import yaml

def run_commands_from_tsv(tsv_file, script_path, script_aclJson_to_excel_py_path, filter_value=None, config=None):
  """
  tsv_file:   パラメータのTSVファイルへのパス
  script_path: 実行するスクリプトへのパス
  filter_value: 起動時に受け取った引数(1列目と比較する値)
         Noneの場合はすべての行を対象とする
  config: 認証情報を含む設定辞書
  """
  if not os.path.exists(tsv_file):
    print(f"TSVファイルが見つかりません: {tsv_file}")
    sys.exit(1)

  if not os.path.exists(script_path):
    print(f"スクリプトファイルが見つかりません: {script_path}")
    sys.exit(1)

  with open(tsv_file, 'r', newline='', encoding='utf-8') as f:
    reader = csv.reader(f, delimiter='\t')
    for idx, row in enumerate(reader, start=1):
      # TSVの行のカラム数チェック (最低でも2列必要)
      if len(row) < 2:
        print(f"行 {idx} が不正です。スキップします。")
        print(f"  row:{row}")
        continue

      # filter_value が None でないときは、1列目を比較してスキップ判定
      if filter_value is not None and row[0].strip() != filter_value:
        continue

      arg1, arg2 = row[0].strip(), row[1].strip()  # arg1: appid, arg2: api_token
      cmd = ['python', script_path, arg1, arg2]  # 最初にappidとapi_tokenを渡す
      if config:
        # 認証情報を追加（subdomain, username, password）
        cmd.extend([config['subdomain'], config['username'], config['password']])
      # arg2のマスク処理
      if len(arg2) <= 2:
        masked_arg2 = arg2  # 2文字以下はそのまま
      else:
        # api_tokenは中間を*に置換
        masked_arg2 = arg2[:2] + '*' * (len(arg2)-4) + arg2[-2:]
      cmd_for_print = ['python', script_path, arg1, masked_arg2]
      if config:
        # パスワードのマスク処理
        masked_config = config.copy()
        password = config['password']
        if len(password) <= 2:
          masked_config['password'] = password
        else:
          # パスワードは中間を*に置換
          masked_config['password'] = password[:2] + '*' * (len(password)-4) + password[-2:]
        cmd_for_print.extend([config['subdomain'], config['username'], masked_config['password']])
      print(f"実行中(1): {' '.join(cmd_for_print)}")

      try:
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print(f"成功: 行 {idx}")
        print(result.stdout)
      except subprocess.CalledProcessError as e:
        print(f"エラー: 行 {idx} のコマンドが失敗しました。")
        print(e.stderr)
     
      # aclJson_to_excel.pyの実行（認証情報は不要）
      cmd = ['python', script_aclJson_to_excel_py_path, arg1]
      print(f"実行中(2): {' '.join(cmd)}")
      try:
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print(f"成功: 行 {idx}")
        print(result.stdout)
      except subprocess.CalledProcessError as e:
        print(f"エラー: 行 {idx} のコマンドが失敗しました。 {script_aclJson_to_excel_py_path}")
        print(e.stderr)

def load_config(config_path='config_UserAccount.yaml'):
  """設定ファイルを読み込む"""
  try:
    with open(config_path, 'r', encoding='utf-8') as f:
      return yaml.safe_load(f)
  except FileNotFoundError:
    print(f"警告: 設定ファイル {config_path} が見つかりません。")
    return {}

if __name__ == '__main__':
  """
  実行方法:
   python run_scripts.py -a                    # 全部の行を実行 (デフォルト設定ファイルを使用)
   python run_scripts.py -all                  # 全部の行を実行 (デフォルト設定ファイルを使用)
   python run_scripts.py <値>                  # 1列目と一致する行のみ実行 (デフォルト設定ファイルを使用)
   python run_scripts.py -c <設定ファイル> <値>  # 指定した設定ファイルを使用
   python run_scripts.py -d <サブドメイン> <ユーザー名> <パスワード> <値>  # 認証情報を直接指定
  """

  # 引数の解析
  if len(sys.argv) < 2:
    print("使い方:")
    print(" python run_scripts.py -a                    # 全部の行を実行")
    print(" python run_scripts.py -all                  # 全部の行を実行")
    print(" python run_scripts.py <値>                  # 1列目と一致する行のみ実行")
    print(" python run_scripts.py -c <設定ファイル> <値>  # 指定した設定ファイルを使用")
    print(" python run_scripts.py -d <サブドメイン> <ユーザー名> <パスワード> <値>  # 認証情報を直接指定")
    sys.exit(0)

  # TSVファイルと実行するスクリプトのパスを指定
  tsv_file = 'run_scripts_params.tsv'
  script_path = './download2yaml_excel.py'
  script_aclJson_to_excel_py_path = 'aclJson_to_excel.py'

  # 引数の処理
  config = None
  filter_value = None

  if sys.argv[1] == '-c':
    # 設定ファイルを指定する場合
    if len(sys.argv) < 4:
      print("エラー: 設定ファイルと値を指定してください")
      sys.exit(1)
    config = load_config(sys.argv[2])
    filter_value = sys.argv[3] if sys.argv[3] not in ('-a', '-all') else None
  elif sys.argv[1] == '-d':
    # 認証情報を直接指定する場合
    if len(sys.argv) < 6:
      print("エラー: サブドメイン、ユーザー名、パスワード、値を指定してください")
      sys.exit(1)
    config = {
      'subdomain': sys.argv[2],
      'username': sys.argv[3],
      'password': sys.argv[4]
    }
    filter_value = sys.argv[5] if sys.argv[5] not in ('-a', '-all') else None
  else:
    # デフォルトの設定ファイルを使用する場合
    config = load_config()
    if sys.argv[1] in ('-a', '-all'):
      filter_value = None
    else:
      filter_value = sys.argv[1]

  # 設定が取得できているか確認
  if not config or not all(key in config for key in ['subdomain', 'username', 'password']):
    print("エラー: 認証情報が不足しています")
    sys.exit(1)

  # コマンドを実行
  run_commands_from_tsv(tsv_file, script_path, script_aclJson_to_excel_py_path, filter_value, config)
