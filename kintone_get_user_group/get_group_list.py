import argparse
import sys
import base64
from getpass import getpass
from typing import List, Dict, Any
import os
import requests
import yaml
import logging

class ArgumentParser:
  @staticmethod
  def parse_arguments():
    parser = argparse.ArgumentParser(
      description='Kintoneの全グループの名前とIDの対照表をYAMLファイルとして出力します。\n' +
                  '引数が指定されていない場合、config_UserAccount.yamlから認証情報を読み込みます。',
      formatter_class=argparse.RawTextHelpFormatter
    )
    # 認証情報用の引数
    auth_group = parser.add_argument_group('認証オプション')
    auth_group.add_argument('--subdomain', help='Kintoneのサブドメイン (例: sample)')
    auth_group.add_argument('--username', help='管理者ユーザーのログイン名 (例: user@example.com)')
    auth_group.add_argument('--password', help='管理者ユーザーのパスワード')
   
    # その他のオプション
    other_group = parser.add_argument_group('その他のオプション')
    other_group.add_argument('--config', help='認証情報を含む設定ファイルのパス (例: ../get_app/data_config.yaml)')
    other_group.add_argument('--out', default='.', help='出力先ディレクトリのパス (デフォルト: 現在のディレクトリ)')
    other_group.add_argument('--silent', action='store_true', help='サイレントモードを有効にします。詳細なログを表示しません。')
   
    return parser.parse_args()

class KintoneClient:
  def __init__(self, subdomain: str, username: str, password: str, logger: logging.Logger):
    self.subdomain = subdomain
    self.headers = self._get_auth_header(username, password)
    self.logger = logger

  @staticmethod
  def _get_auth_header(username: str, password: str) -> Dict[str, str]:
    credentials = f"{username}:{password}"
    base64_credentials = base64.b64encode(credentials.encode('utf-8')).decode('utf-8')
    return {
      'X-Cybozu-Authorization': base64_credentials
    }

  def _fetch_data(self, endpoint: str, params: Dict[str, Any], key: str) -> List[Dict[str, Any]]:
    url = f"https://{self.subdomain}.cybozu.com/v1/{endpoint}.json"
    data = []
    size = 100
    offset = 0

    while True:
      current_params = params.copy()
      current_params.update({'size': size, 'offset': offset})
      response = requests.get(url, headers=self.headers, params=current_params)
      if response.status_code != 200:
        self.logger.error(f"{endpoint.capitalize()}の取得に失敗しました: {response.status_code} {response.text}")
        sys.exit(1)
      batch = response.json().get(key, [])
      if not batch:
        break
      data.extend(batch)
      if len(batch) < size:
        break
      offset += size
      self.logger.debug(f"Fetched {len(batch)} items from {endpoint} (offset: {offset})")
    self.logger.info(f"全{endpoint}を取得しました。総数: {len(data)}")
    return data

  def get_all_groups(self) -> List[Dict[str, Any]]:
    return self._fetch_data('groups', {}, 'groups')

def setup_logging(silent: bool) -> logging.Logger:
  logger = logging.getLogger("KintoneGroupExporter")
  logger.setLevel(logging.DEBUG if not silent else logging.WARNING)
  handler = logging.StreamHandler(sys.stdout)
  formatter = logging.Formatter('%(levelname)s: %(message)s')
  handler.setFormatter(formatter)
  if not logger.handlers:
    logger.addHandler(handler)
  return logger

def load_config(config_path: str) -> Dict[str, Any]:
  # パスの展開と正規化
  expanded_path = os.path.expanduser(config_path)
  absolute_path = os.path.abspath(expanded_path)

  if not os.path.isfile(absolute_path):
    raise FileNotFoundError(f"設定ファイルが見つかりません: {absolute_path}")
 
  # 拡張子の確認
  _, ext = os.path.splitext(absolute_path)
  if ext.lower() not in ['.yaml', '.yml']:
    raise ValueError(f"設定ファイルの拡張子が無効です: {absolute_path}. 有効な拡張子は .yaml または .yml です。")
 
  with open(absolute_path, 'r', encoding='utf-8') as file:
    try:
      config = yaml.safe_load(file)
      return config
    except yaml.YAMLError as e:
      raise ValueError(f"設定ファイルの解析に失敗しました: {e}")

def main():
  # 引数の解析
  args = ArgumentParser.parse_arguments()
 
  # ロギングの設定
  logger = setup_logging(args.silent)
 
  # 認証情報の初期化
  subdomain = args.subdomain
  username = args.username
  password = args.password

  # 引数が指定されていない場合、デフォルトのconfig_UserAccount.yamlを使用
  if not (subdomain and username and password):
    default_config = 'config_UserAccount.yaml'
    try:
      config = load_config(default_config)
      if not subdomain:
        subdomain = config.get('subdomain')
      if not username:
        username = config.get('username')
      if not password:
        password = config.get('password')
      logger.info(f"デフォルト設定ファイル '{default_config}' から認証情報を読み込みました。")
    except Exception as e:
      logger.error(f"デフォルト設定ファイルの読み込みに失敗しました: {e}")
      sys.exit(1)

  # 必要な認証情報が揃っているか確認
  missing = []
  if not subdomain:
    missing.append('subdomain')
  if not username:
    missing.append('username')
  if not password:
    missing.append('password')
  if missing:
    logger.error(f"以下の認証情報が不足しています: {', '.join(missing)}")
    sys.exit(1)

  # パスワードのプロンプト入力を削除（config_UserAccount.yamlからの読み込みを優先）

  # 出力ディレクトリの確認
  output_dir = os.path.expanduser(args.out)
  output_dir = os.path.abspath(output_dir)
  if not os.path.isdir(output_dir):
    try:
      os.makedirs(output_dir)
      logger.info(f"出力先ディレクトリ '{output_dir}' を作成しました。")
    except Exception as e:
      logger.error(f"出力先ディレクトリ '{output_dir}' の作成に失敗しました: {e}")
      sys.exit(1)

  output_file = os.path.join(output_dir, 'group_list.yaml')

  # Kintoneクライアントの初期化
  logger.info("認証情報を設定中...")
  client = KintoneClient(subdomain, username, password, logger)

  # グループデータの取得
  logger.info("全グループを取得中...")
  all_groups = client.get_all_groups()

  # グループCODEとグループ名の対照表を作成
  group_mapping = {}
  for group in all_groups:
    print(group)
    group_name = group.get('name')
    group_code = group.get('code')
    if group_name and group_code:
      group_mapping[group_code] = group_name # キーと値を入れ替え
      logger.debug(f"グループCODE: {group_code}, グループ名: {group_name}")
    else:
      logger.warning(f"グループ情報が不完全です: {group}")

  logger.info(f"グループの対照表を作成しました。総グループ数: {len(group_mapping)}")

  # YAMLファイルに出力
  try:
    with open(output_file, 'w', encoding='utf-8') as yaml_file:
      yaml.dump(group_mapping, yaml_file, allow_unicode=True, sort_keys=False)
    logger.info(f"YAMLファイル '{output_file}' を作成しました。")
  except Exception as e:
    logger.error(f"YAMLファイル '{output_file}' の作成に失敗しました: {e}")
    sys.exit(1)

if __name__ == "__main__":
  main()
