import argparse
import sys
import base64
from getpass import getpass
from typing import List, Dict, Any
import logging
import yaml
import os
from datetime import datetime

from lib.kintone_userlib.user import User
from lib.kintone_userlib.group import Group
from lib.kintone_userlib.client import KintoneClient

class ArgumentParser:
    @staticmethod
    def parse_arguments():
        parser = argparse.ArgumentParser(
            description='Kintoneの全ユーザーと各ユーザーの所属グループをExcelに出力します。\n\n引数を省略した場合、config_UserAccount.yaml を参照して認証情報を取得し、デフォルトの出力ファイル名を使用します。',
            formatter_class=argparse.RawTextHelpFormatter
        )
        parser.add_argument('--subdomain', help='Kintoneのサブドメイン (例: sample)')
        parser.add_argument('--username', help='管理者ユーザーのログイン名 (例: user@example.com)')
        parser.add_argument('--password', help='管理者ユーザーのパスワード (指定しない場合、プロンプトで入力)')
        parser.add_argument('--output', default='kintone_users_groups.xlsx', help='出力するExcelファイルの名前 (デフォルト: kintone_users_groups.xlsx)')
        parser.add_argument('--silent', action='store_true', help='サイレントモードを有効にします。詳細なログを表示しません。')
        
        return parser.parse_args()

class DataProcessor:
    def __init__(self, users: List[User], groups: List[Group], logger: logging.Logger):
        self.users = users
        self.groups = groups
        self.logger = logger
        self.user_id_to_info: Dict[str, Dict[str, Any]] = {}

    def filter_groups(self) -> List[Group]:
        filtered = [group for group in self.groups if group.name != 'Everyone' and group.code]
        self.logger.info(f"「Everyone」を除外し、codeが存在するグループ数: {len(filtered)}")
        return filtered

    def organize_groups(self, filtered_groups: List[Group]) -> List[str]:
        group_names = [group.name for group in filtered_groups]
        self.logger.info(f"グループ名一覧: {group_names}")

        if 'Administrators' in group_names:
            group_names.remove('Administrators')
            group_names.insert(0, 'Administrators')
            self.logger.info("「Administrators」グループをグループ名一覧の最初に配置しました。")
        else:
            self.logger.info("「Administrators」グループがグループ名一覧に存在しませんでした。")
        self.logger.info(f"最終的なグループ名一覧: {group_names}")
        return group_names

    def map_users(self):
        for user in self.users:
            user_id = str(user.id)
            self.user_id_to_info[user_id] = {
                'ユーザーID': user_id,
                'ステータス': '停止中' if not user.is_active else '',
                'ログイン名': user.username,
                '氏名': user.name,
                'メールアドレス': user.email,
                '所属グループ一覧': [group.name for group in user.groups],
                '最終アクセス日': user.last_login.strftime('%Y-%m-%d %H:%M:%S') if user.last_login else '',
                '経過日数': user.days_since_last_login() if user.last_login else ''
            }
        self.logger.info(f"ユーザー情報をマッピングしました。総ユーザー数: {len(self.user_id_to_info)}")

    def export_group_user_list(self, filtered_groups: List[Group]):
        """グループとユーザーの関連をYAMLファイルとして出力"""
        self.logger.info("group_user_list.yaml、group_user_list_NoUse.yaml、user_list.yaml、group_user_raw_list.yaml を生成中...")
        
        active_group_data = {}
        inactive_group_data = {}
        raw_group_data = {}
        
        # ユーザーリスト用のデータ構造を準備
        user_list_data = {}
        
        # まず全グループの基本情報を設定
        for group in filtered_groups:
            group_code = group.code
            active_group_data[group_code] = {
                'name': group.name,
                'users': []
            }
            inactive_group_data[group_code] = {
                'name': group.name,
                'users': []
            }
            
            # グループ内のユーザー情報を取得
            for user in group.users:
                user_info = {
                    'username': user.username,
                    'email': user.email,
                    'id': str(user.id),
                    'isDisabled': not user.is_active
                }
                
                # ユーザーリストデータにも追加
                if user.username not in user_list_data:
                    user_list_data[user.username] = {
                        'code': user.username,
                        'username': user.username,
                        'name': user.name,
                        'email': user.email,
                        'valid': user.is_active,
                        'isDisabled': not user.is_active
                    }
                
                # ユーザーの状態を確認
                if user.is_active:
                    active_group_data[group_code]['users'].append(user_info)
                else:
                    inactive_group_data[group_code]['users'].append(user_info)
                
                # raw_group_data にも追加
                if group_code not in raw_group_data:
                    raw_group_data[group_code] = {
                        'name': group.name,
                        'users': []
                    }
                raw_group_data[group_code]['users'].append(user_info)
        
        # Everyoneグループを追加
        active_everyone_users = []
        inactive_everyone_users = []
        for user in self.users:
            user_info = {
                'username': user.username,
                'email': user.email,
                'id': str(user.id),
                'isDisabled': not user.is_active
            }
            
            # ユーザーリストデータにも追加
            if user.username not in user_list_data:
                user_list_data[user.username] = {
                    'code': user.username,
                    'username': user.username,
                    'name': user.name,
                    'email': user.email,
                    'valid': user.is_active,
                    'isDisabled': not user.is_active
                }
            
            if user.is_active:
                active_everyone_users.append(user_info)
            else:
                inactive_everyone_users.append(user_info)
        
        active_group_data['everyone'] = {
            'name': 'Everyone',
            'users': active_everyone_users
        }
        
        inactive_group_data['everyone'] = {
            'name': 'Everyone',
            'users': inactive_everyone_users
        }
        
        # 空のグループを削除
        active_group_data = {k: v for k, v in active_group_data.items() if v['users']}
        inactive_group_data = {k: v for k, v in inactive_group_data.items() if v['users']}
        
        # YAMLファイルに出力
        try:
            # アクティブユーザー用のファイル
            with open('group_user_list.yaml', 'w', encoding='utf-8') as f:
                yaml.dump(active_group_data, f, allow_unicode=True, sort_keys=False)
            
            # 停止中ユーザー用のファイル
            with open('group_user_list_NoUse.yaml', 'w', encoding='utf-8') as f:
                yaml.dump(inactive_group_data, f, allow_unicode=True, sort_keys=False)
            
            # ユーザーリストファイルを出力
            with open('user_list.yaml', 'w', encoding='utf-8') as f:
                yaml.dump(user_list_data, f, allow_unicode=True, sort_keys=False)
            
            # rawユーザー用のファイル
            with open('group_user_raw_list.yaml', 'w', encoding='utf-8') as f:
                yaml.dump(raw_group_data, f, allow_unicode=True, sort_keys=False)
            
            self.logger.info("group_user_list.yaml、group_user_list_NoUse.yaml、user_list.yaml、group_user_raw_list.yaml の生成が完了しました。")
        except Exception as e:
            self.logger.error(f"YAMLファイルの生成中にエラーが発生しました: {e}")

def setup_logging(silent: bool, debug: bool) -> logging.Logger:
    logger = logging.getLogger("KintoneExporter")
    # 既存のハンドラーをクリア
    logger.handlers.clear()
    
    if debug:
        log_level = logging.DEBUG
    elif silent:
        log_level = logging.WARNING
    else:
        log_level = logging.INFO
        
    logger.setLevel(log_level)
    handler = logging.StreamHandler(sys.stdout)
    handler.setLevel(log_level)  # ハンドラーのレベルも設定
    formatter = logging.Formatter('%(levelname)s: %(message)s')
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    
    return logger

def load_config(config_path: str) -> Dict[str, Any]:
    with open(config_path, 'r', encoding='utf-8') as file:
        return yaml.safe_load(file)

def main():
    args = ArgumentParser.parse_arguments()
    logger = setup_logging(args.silent, False)

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

    # Kintoneクライアントの初期化
    logger.info("認証情報を設定中...")
    client = KintoneClient(subdomain, username, password, logger)

    # データの取得
    logger.info("全ユーザーを取得中...")
    all_users = client.get_all_users()

    logger.info("全グループを取得中...")
    all_groups = client.get_all_groups()

    # データの処理
    processor = DataProcessor(all_users, all_groups, logger)
    processor.map_users()
    filtered_groups = processor.filter_groups()
    group_names = processor.organize_groups(filtered_groups)
    processor.export_group_user_list(filtered_groups)

    logger.info("処理が完了しました。")

if __name__ == "__main__":
    main() 