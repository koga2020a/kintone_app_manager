import argparse
import sys
import base64
from getpass import getpass
import logging
import yaml
import requests

class KintoneClient:
    def __init__(self, subdomain: str, username: str, password: str, logger: logging.Logger):
        self.subdomain = subdomain
        self.headers = self._get_auth_header(username, password)
        self.logger = logger
        self.base_url = f"https://{subdomain}.cybozu.com"

    @staticmethod
    def _get_auth_header(username: str, password: str) -> dict:
        credentials = f"{username}:{password}"
        base64_credentials = base64.b64encode(credentials.encode('utf-8')).decode('utf-8')
        return {
            'X-Cybozu-Authorization': base64_credentials
        }

    def _make_request(self, method: str, url: str, **kwargs) -> requests.Response:
        try:
            response = requests.request(method, url, headers=self.headers, **kwargs)
            if response.status_code != 200:
                self.logger.error("=== エラーレスポンス ===")
                self.logger.error(f"ステータスコード: {response.status_code}")
                self.logger.error(f"レスポンスヘッダー: {dict(response.headers)}")
                self.logger.error(f"レスポンスボディ: {response.text}")
                self.logger.error("=====================")
                sys.exit(1)
            return response
        except requests.exceptions.RequestException as e:
            self.logger.error("=== リクエスト情報 ===")
            self.logger.error(f"メソッド: {method}")
            self.logger.error(f"URL: {url}")
            self.logger.error(f"ヘッダー: {self.headers}")
            self.logger.error(f"パラメータ: {kwargs}")
            self.logger.error("=== エラー内容 ===")
            self.logger.error(f"リクエスト中に例外が発生しました: {e}")
            self.logger.error("=================")
            sys.exit(1)

    def get_all_groups(self) -> list:
        """グループ一覧を取得"""
        url = f"{self.base_url}/v1/groups.json"
        
        # リクエスト内容をデバッグログ出力
        self.logger.debug("=== グループ一覧取得 リクエスト ===")
        self.logger.debug(f"リクエストURL: {url}")
        self.logger.debug(f"リクエストヘッダー: {self.headers}")
        self.logger.debug(f"リクエストメソッド: GET")
        self.logger.debug("================================")
        
        response = self._make_request("GET", url)
        
        # システム固定グループを除外
        SYSTEM_GROUPS = {'Administrators', 'everyone'}
        groups = response.json().get('groups', [])
        filtered_groups = [group for group in groups if group['code'] not in SYSTEM_GROUPS]
        
        return filtered_groups

    def search_users(self, keyword: str) -> list:
        """ユーザーを検索"""
        all_users = []
        offset = 0
        limit = 100  # Kintone APIの1回のリクエスト上限

        while True:
            url = f"{self.base_url}/v1/users.json"
            params = {
                'offset': offset,
                'size': limit
            }
            self.logger.debug(f"ユーザー検索用URL: {url}, パラメータ: {params}")
            response = self._make_request("GET", url, params=params)
            
            users = response.json().get('users', [])
            if not users:
                break
                
            all_users.extend(users)
            if len(users) < limit:
                break
                
            offset += limit

        if not keyword:
            return all_users

        # 検索キーワードに一致するユーザーをフィルタリング
        matched_users = []
        keyword = keyword.lower()
        for user in all_users:
            code = user.get('code', '') or ''
            name = user.get('name', '') or ''
            email = user.get('email', '') or ''
            if (keyword in code.lower() or
                keyword in name.lower() or
                keyword in email.lower()):
                matched_users.append(user)
        return matched_users

    def get_user_groups(self, user_code: str) -> list:
        """ユーザーの所属グループを取得"""
        url = f"{self.base_url}/v1/user/groups.json"
        params = {'code': user_code}
        
        # リクエスト内容をデバッグログ出力
        self.logger.debug("=== ユーザーのグループ情報取得 リクエスト ===")
        self.logger.debug(f"リクエストURL: {url}")
        self.logger.debug(f"リクエストヘッダー: {self.headers}")
        self.logger.debug(f"リクエストパラメータ: {params}")
        self.logger.debug(f"リクエストメソッド: GET")
        self.logger.debug("================================")
        
        response = self._make_request("GET", url, params=params)
        return response.json().get('groups', [])

    def get_group_users(self, group_code: str) -> list:
        """グループの現在のユーザー一覧を取得"""
        url = f"{self.base_url}/v1/group/users.json"
        params = {'code': group_code}
        
        # リクエスト内容をデバッグログ出力
        self.logger.debug("=== グループのユーザー一覧取得 リクエスト ===")
        self.logger.debug(f"リクエストURL: {url}")
        self.logger.debug(f"リクエストヘッダー: {self.headers}")
        self.logger.debug(f"リクエストパラメータ: {params}")
        self.logger.debug(f"リクエストメソッド: GET")
        self.logger.debug("================================")
        
        response = self._make_request("GET", url, params=params)
        return response.json().get('users', [])

    def add_user_to_group(self, group_code: str, user_codes: list) -> bool:
        """ユーザーをグループに追加（既存のユーザーを保持）"""
        url = f"{self.base_url}/v1/group/users.json"
        
        # 現在のグループユーザーを取得
        existing_users = self.get_group_users(group_code)
        existing_user_codes = [user['code'] for user in existing_users]
        
        # 新しいユーザーを追加（重複を避ける）
        updated_user_codes = list(set(existing_user_codes + user_codes))
        
        data = {
            "code": group_code,
            "users": updated_user_codes  # 既存と新規のユーザーを統合
        }
        
        # リクエスト内容をデバッグログ出力
        self.logger.debug("=== グループユーザー追加 リクエスト ===")
        self.logger.debug(f"リクエストURL: {url}")
        self.logger.debug(f"リクエストヘッダー: {self.headers}")
        self.logger.debug(f"リクエストボディ: {data}")
        self.logger.debug("================================")
        
        self._make_request("PUT", url, json=data)
        return True

    def remove_user_from_group(self, group_code: str, user_codes: list) -> bool:
        """ユーザーをグループから削除（既存の他のユーザーを保持）"""
        url = f"{self.base_url}/v1/group/users.json"
        
        # 現在のグループユーザーを取得
        existing_users = self.get_group_users(group_code)
        existing_user_codes = [user['code'] for user in existing_users]
        
        # 削除するユーザーを除外
        updated_user_codes = [code for code in existing_user_codes if code not in user_codes]
        
        data = {
            "code": group_code,
            "users": updated_user_codes  # 削除後のユーザーリストを送信
        }
        
        # リクエスト内容をデバッグログ出力
        self.logger.debug("=== グループユーザー削除 リクエスト ===")
        self.logger.debug(f"リクエストURL: {url}")
        self.logger.debug(f"リクエストヘッダー: {self.headers}")
        self.logger.debug(f"リクエストボディ: {data}")
        self.logger.debug("================================")
        
        self._make_request("PUT", url, json=data)
        return True

    def get_group_by_name_or_code(self, group_identifier: str) -> dict:
        """グループ名またはコードからグループを検索"""
        groups = self.get_all_groups()
        for group in groups:
            if group['code'] == group_identifier or group['name'] == group_identifier:
                return group
        return None

class GroupManager:
    def __init__(self, client: KintoneClient, logger: logging.Logger):
        self.client = client
        self.logger = logger

    def list_groups(self):
        """グループ一覧を表示"""
        groups = self.client.get_all_groups()
        if not groups:
            self.logger.info("グループが見つかりませんでした")
            return
        
        print("\n利用可能なグループ:")
        print("-" * 50)
        for i, group in enumerate(groups, 1):
            print(f"{i}. {group['name']} (コード: {group['code']})")
        print("-" * 50)
        return groups

    def search_users(self, keyword: str):
        """ユーザーを検索して表示し、選択されたユーザーのグループ情報を表示"""
        users = self.client.search_users(keyword)
        if not users:
            print(f"'{keyword}' に一致するユーザーは見つかりませんでした")
            return

        print("\n検索結果:")
        print("-" * 80)
        print(f"{'No.':<4} {'ログイン名':<30} {'表示名':<30} {'メールアドレス':<30}")
        print("-" * 80)
        
        for i, user in enumerate(users, 1):
            username = user.get('code', '')
            name = user.get('name', '')
            email = user.get('email', '')
            print(f"{i:<4} {username:<30} {name:<30} {email:<30}")
        
        print("-" * 80)
        print(f"合計: {len(users)}件")

        # 1名の場合は自動的に選択
        selected_user = users[0] if len(users) == 1 else None

        # 複数名の場合は選択を促す
        if len(users) > 1:
            while True:
                try:
                    choice = int(input("\n詳細を表示するユーザーの番号を入力してください (0でキャンセル): "))
                    if choice == 0:
                        return
                    if 1 <= choice <= len(users):
                        selected_user = users[choice - 1]
                        break
                    print("無効な選択です。もう一度入力してください。")
                except ValueError:
                    print("数字を入力してください。")

        # 選択されたユーザーのグループ情報を表示
        if selected_user:
            user_code = selected_user['code']
            groups = self.client.get_user_groups(user_code)
            
            print(f"\n{selected_user['name']} ({user_code}) の所属グループ:")
            print("-" * 50)
            if groups:
                for i, group in enumerate(groups, 1):
                    if group['code'] not in {'Administrators', 'everyone'}:
                        print(f"{i}. {group['name']} (コード: {group['code']})")
            else:
                print("所属グループはありません")
            print("-" * 50)

    def set_user_group(self, user_code: str, group_identifier: str = None):
        """ユーザーのグループを設定"""
        if group_identifier is None:
            # 対話モード
            groups = self.list_groups()
            if not groups:
                return
            
            while True:
                try:
                    choice = int(input("\n設定するグループの番号を入力してください (0でキャンセル): "))
                    if choice == 0:
                        return
                    if 1 <= choice <= len(groups):
                        group_code = groups[choice - 1]['code']
                        break
                    print("無効な選択です。もう一度入力してください。")
                except ValueError:
                    print("数字を入力してください。")
        else:
            # グループ名またはコードで検索
            group = self.client.get_group_by_name_or_code(group_identifier)
            if not group:
                self.logger.error(f"グループ '{group_identifier}' が見つかりません")
                sys.exit(1)
            group_code = group['code']
        
        # 現在のグループを取得
        current_groups = self.client.get_user_groups(user_code)
        current_group_codes = [g['code'] for g in current_groups 
                              if g['code'] not in {'Administrators', 'everyone'}]
        
        # 現在のグループから削除（システムグループを除外）
        for curr_group_code in current_group_codes:
            if curr_group_code != group_code:
                self.logger.info(f"ユーザー {user_code} をグループ {curr_group_code} から削除します")
                self.client.remove_user_from_group(curr_group_code, [user_code])
        
        # 新しいグループに追加
        if group_code not in current_group_codes:
            self.logger.info(f"ユーザー {user_code} をグループ {group_code} に追加します")
            self.client.add_user_to_group(group_code, [user_code])
            self.logger.info("グループ設定を更新しました")
        else:
            self.logger.info(f"ユーザー {user_code} は既にグループ {group_code} に所属しています")

def setup_logging(silent: bool = False, debug: bool = False) -> logging.Logger:
    logger = logging.getLogger("KintoneGroupManager")
    if debug:
        logger.setLevel(logging.DEBUG)
    elif silent:
        logger.setLevel(logging.WARNING)
    else:
        logger.setLevel(logging.INFO)
    handler = logging.StreamHandler(sys.stdout)
    if debug:
        formatter = logging.Formatter('%(levelname)s: %(message)s')
    else:
        formatter = logging.Formatter('%(levelname)s: %(message)s')
    handler.setFormatter(formatter)
    if not logger.handlers:
        logger.addHandler(handler)
    return logger

def load_config(config_path: str = 'config_UserAccount.yaml') -> dict:
    try:
        with open(config_path, 'r', encoding='utf-8') as file:
            return yaml.safe_load(file)
    except FileNotFoundError:
        # config_UserAccount.yamlが見つからない場合は.kintone.envを試す
        try:
            with open('.kintone.env', 'r', encoding='utf-8') as file:
                return yaml.safe_load(file)
        except Exception as e:
            print(f".kintone.envの読み込みにも失敗しました: {e}")
            sys.exit(1)
    except Exception as e:
        print(f"設定ファイルの読み込みに失敗しました: {e}")
        sys.exit(1)

def main():
    parser = argparse.ArgumentParser(
        description='kintoneグループ管理ツール',
        formatter_class=argparse.RawTextHelpFormatter
    )
    parser.add_argument('command', nargs='?', help='コマンド (set/list) または検索キーワード')
    parser.add_argument('user', nargs='?', help='ユーザー名 (setコマンド用)')
    parser.add_argument('group', nargs='?', help='グループ名 (setコマンド用)')
    parser.add_argument('--config', default='config_UserAccount.yaml', help='設定ファイルのパス')
    parser.add_argument('--silent', action='store_true', help='詳細なログを表示しない')
    parser.add_argument('--debug', action='store_true', help='デバッグログを表示する')
    parser.add_argument('--search', action='store_true', help='ユーザー検索モードを有効にする')

    args = parser.parse_args()
    
    if not args.command:
        parser.print_help()
        return

    # ログと設定の初期化
    logger = setup_logging(args.silent, args.debug)
    config = load_config(args.config)

    # クライアントの初期化
    client = KintoneClient(
        config['subdomain'],
        config['username'],
        config['password'],
        logger
    )

    # マネージャーの初期化
    manager = GroupManager(client, logger)

    # --search オプションが指定されている場合は、commandをキーワードとして検索
    if args.search:
        manager.search_users(args.command)
        return

    # コマンドの処理
    if args.command == 'list':
        manager.list_groups()
    elif args.command == 'set':
        if not args.user:
            print("エラー: setコマンドにはユーザー名が必要です")
            sys.exit(1)
        manager.set_user_group(args.user, args.group)
    elif args.command.startswith('--'):
        # スイッチオプションの場合は何もしない
        pass
    else:
        # その他の場合は検索キーワードとして処理
        manager.search_users(args.command)

if __name__ == "__main__":
    main()