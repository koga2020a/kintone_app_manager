#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Kintone関連ツールの統合実行スクリプト

このスクリプトは.kintone.envファイルから設定を読み込み、
kintone_get_user_group、kintone_get_appjson、kintone_group_cliの
機能を連携して実行するためのものです。
"""

import os
import sys
import yaml
import argparse
import subprocess
import logging
from pathlib import Path
from datetime import datetime

# 定数定義
SCRIPT_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = SCRIPT_DIR / "output"
PREVIOUS_OUTPUT_DIR = SCRIPT_DIR / "previous_output"
BACKUP_DIR = SCRIPT_DIR / "backup"
ENV_FILE = SCRIPT_DIR / ".kintone.env"
CONFIG_FILE = SCRIPT_DIR / "config_UserAccount.yaml"

# 各ディレクトリのパス
USER_GROUP_DIR = SCRIPT_DIR / "kintone_get_user_group"
APPJSON_DIR = SCRIPT_DIR / "kintone_get_appjson"
GROUP_CLI_DIR = SCRIPT_DIR / "kintone_group_cli"

# 出力ファイル情報定義
OUTPUT_FILE_INFO = {
    "excel": [
        {
            "name": "kintone_users_groups_[日時].xlsx",
            "description": "ユーザーとグループの一覧情報",
            "command": "users",
            "args": "--format excel (デフォルト)"
        },
        {
            "name": "acl_report_[アプリID]_[日時].xlsx",
            "description": "アプリのACL情報（ユーザー名・グループ名を反映）",
            "command": "acl",
            "args": "--id [アプリID] (省略時は全アプリ対象)"
        }
    ],
    "csv": [
        {
            "name": "kintone_users_groups_[日時].csv",
            "description": "ユーザーとグループの一覧情報（CSV形式）",
            "command": "users",
            "args": "--format csv"
        },
        {
            "name": "[アプリID]permission_target_user_names.csv",
            "description": "アプリに出現するユニークなユーザー名一覧",
            "command": "acl",
            "args": "--id [アプリID] (自動生成される補助ファイル)"
        }
    ],
    "tsv": [
        {
            "name": "(現在TSV形式の出力はサポートされていません)",
            "description": "",
            "command": "",
            "args": ""
        }
    ]
}

# ログ設定
def setup_logging():
    """ロギングの設定"""
    log_dir = SCRIPT_DIR / "logs"
    log_dir.mkdir(exist_ok=True)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = log_dir / f"kintone_runner_{timestamp}.log"
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()
        ]
    )
    
    return logging.getLogger("kintone_runner")

# 設定ファイルの読み込み
def load_env_config(env_file=None):
    """
    .kintone.env ファイルを読み込み、設定情報を返す
    """
    if env_file is None:
        env_file = ENV_FILE
    
    if not env_file.exists():
        print(f"エラー: 設定ファイル {env_file} が見つかりません。")
        sys.exit(1)
    
    try:
        with open(env_file, 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f)
        
        # 必須項目をチェック
        required_keys = ['subdomain', 'username', 'password']
        missing_keys = [key for key in required_keys if key not in config]
        
        if missing_keys:
            print(f"エラー: 設定ファイルに以下の必須項目がありません: {', '.join(missing_keys)}")
            sys.exit(1)
            
        return config
    except Exception as e:
        print(f"エラー: 設定ファイルの読み込み中にエラーが発生しました: {e}")
        sys.exit(1)

# 設定ファイルの作成
def create_config_file(config, config_path=None):
    """
    .kintone.env の内容から config_UserAccount.yaml を作成
    """
    if config_path is None:
        config_path = CONFIG_FILE
    
    try:
        with open(config_path, 'w', encoding='utf-8') as f:
            yaml.dump(config, f, default_flow_style=False)
        return True
    except Exception as e:
        print(f"エラー: config_UserAccount.yaml の作成中にエラーが発生しました: {e}")
        return False

# 出力ファイル情報の表示
def display_output_info():
    """
    生成されるExcel、CSV、TSVファイルの情報を表示
    """
    print("=== Kintone Runner が生成するファイル一覧 ===")
    print("※ JSON、YAMLファイルは除く\n")
    
    for file_type, files in OUTPUT_FILE_INFO.items():
        print(f"【{file_type.upper()}ファイル】")
        for file_info in files:
            if file_info["name"] and file_info["command"]:
                print(f"■ {file_info['name']}")
                print(f"  内容: {file_info['description']}")
                print(f"  コマンド: {file_info['command']} {file_info['args']}")
                print()
            else:
                print(f"■ {file_info['name']}")
                print()
    
    print("※ すべてのファイルは 'all' コマンドでも一括生成できます。")
    print("※ 出力先ディレクトリ: ./output/")

# ユーザーとグループ情報の取得
def get_user_group_info(config, logger, output_format="excel"):
    """
    kintone_get_user_group の機能を呼び出してユーザーとグループ情報を取得
    """
    logger.info("ユーザーとグループ情報の取得を開始します")
    
    script_path = USER_GROUP_DIR / "get_user_group.py"
    
    if not script_path.exists():
        logger.error(f"スクリプトファイルが見つかりません: {script_path}")
        return False
    
    # 出力ディレクトリが存在しない場合は作成
    OUTPUT_DIR.mkdir(exist_ok=True)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = OUTPUT_DIR / f"kintone_users_groups_{timestamp}.xlsx"
    
    cmd = [
        sys.executable, 
        str(script_path),
        "--subdomain", config["subdomain"],
        "--username", config["username"],
        "--password", config["password"],
        "--output", str(output_file)
    ]
    
    try:
        logger.info(f"実行コマンド: {' '.join(cmd).replace(config['password'], '********')}")
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        logger.info(f"ユーザーとグループ情報を {output_file} に出力しました")
        logger.debug(f"出力: {result.stdout}")
        return str(output_file)
    except subprocess.CalledProcessError as e:
        logger.error(f"ユーザーとグループ情報の取得中にエラーが発生しました: {e}")
        logger.error(f"標準出力: {e.stdout}")
        logger.error(f"標準エラー: {e.stderr}")
        return False

# アプリのJSONデータ取得
def get_app_json(config, logger, app_id=None):
    """
    kintone_get_appjson の機能を呼び出してアプリのJSONデータを取得
    """
    logger.info("アプリのJSONデータ取得を開始します")
    
    script_path = APPJSON_DIR / "download2yaml_excel.py"
    
    if not script_path.exists():
        logger.error(f"スクリプトファイルが見つかりません: {script_path}")
        return False
    
    # 出力ディレクトリが存在しない場合は作成
    OUTPUT_DIR.mkdir(exist_ok=True)
    
    # app_tokensからアプリIDとAPIトークンを取得
    app_tokens = config.get('app_tokens', {})
    
    if app_id:
        # 特定のアプリIDが指定された場合
        if str(app_id) not in app_tokens:
            logger.error(f"アプリID {app_id} のAPIトークンが設定されていません")
            return False
            
        api_token = app_tokens[str(app_id)]
        cmd = [
            sys.executable,
            str(script_path),
            str(app_id),
            api_token,
            config["subdomain"],
            config["username"],
            config["password"]
        ]
        
        try:
            logger.info(f"実行コマンド: python {script_path} {app_id} ****** {config['subdomain']} {config['username']} ********")
            result = subprocess.run(cmd, check=True, capture_output=True, text=True)
            logger.info(f"アプリID {app_id} のJSONデータを取得しました")
            logger.debug(f"出力: {result.stdout}")
            return True
        except subprocess.CalledProcessError as e:
            logger.error(f"アプリのJSONデータ取得中にエラーが発生しました: {e}")
            logger.error(f"標準出力: {e.stdout}")
            logger.error(f"標準エラー: {e.stderr}")
            return False
    else:
        # 全てのアプリを処理
        success = True
        for app_id, api_token in app_tokens.items():
            logger.info(f"アプリID {app_id} の処理を開始します")
            cmd = [
                sys.executable,
                str(script_path),
                str(app_id),
                api_token,
                config["subdomain"],
                config["username"],
                config["password"]
            ]
            
            try:
                logger.info(f"実行コマンド: python {script_path} {app_id} ****** {config['subdomain']} {config['username']} ********")
                result = subprocess.run(cmd, check=True, capture_output=True, text=True)
                logger.info(f"アプリID {app_id} のJSONデータを取得しました")
                logger.debug(f"出力: {result.stdout}")
            except subprocess.CalledProcessError as e:
                logger.error(f"アプリID {app_id} のJSONデータ取得中にエラーが発生しました: {e}")
                logger.error(f"標準出力: {e.stdout}")
                logger.error(f"標準エラー: {e.stderr}")
                success = False
                
        return success

# グループ操作
def manage_groups(config, logger, action, params=None):
    """
    kintone_group_cli の機能を呼び出してグループを操作
    
    action: 'list', 'search', 'add', 'remove'
    params: アクションに応じたパラメータ
    """
    logger.info(f"グループ操作 '{action}' を開始します")
    
    script_path = GROUP_CLI_DIR / "group_cli.py"
    
    if not script_path.exists():
        logger.error(f"スクリプトファイルが見つかりません: {script_path}")
        return False
    
    # configファイルが必要なので、一時的に作成
    tmp_config_file = GROUP_CLI_DIR / "config_UserAccount.yaml"
    if not create_config_file(config, tmp_config_file):
        logger.error("一時設定ファイルの作成に失敗しました")
        return False
    
    cmd = [sys.executable, str(script_path)]
    
    # アクションに応じてコマンドラインを構築
    if action == 'list':
        cmd.append('list')
    elif action == 'search':
        if not params or 'keyword' not in params:
            logger.error("検索にはキーワードが必要です")
            return False
        cmd.append('--search')
        cmd.append(params['keyword'])
    elif action == 'add':
        if not params or 'user' not in params or 'group' not in params:
            logger.error("ユーザー追加にはユーザーコードとグループ名/コードが必要です")
            return False
        cmd.extend(['set', params['user'], params['group']])
    elif action == 'remove':
        if not params or 'user' not in params:
            logger.error("ユーザー削除にはユーザーコードが必要です")
            return False
        cmd.extend(['set', params['user']])
    else:
        logger.error(f"不明なアクション: {action}")
        return False
    
    try:
        logger.info(f"実行コマンド: {' '.join(cmd)}")
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        logger.info(f"グループ操作 '{action}' が完了しました")
        logger.info(f"出力: {result.stdout}")
        
        # 一時ファイルを削除
        if tmp_config_file.exists():
            tmp_config_file.unlink()
            
        return result.stdout
    except subprocess.CalledProcessError as e:
        logger.error(f"グループ操作中にエラーが発生しました: {e}")
        logger.error(f"標準出力: {e.stdout}")
        logger.error(f"標準エラー: {e.stderr}")
        
        # 一時ファイルを削除
        if tmp_config_file.exists():
            tmp_config_file.unlink()
            
        return False

# 既存のディレクトリを探す関数
def find_existing_directory(base_dir, app_id):
    """
    指定されたディレクトリ内で、特定のアプリIDで始まるディレクトリを探す

    Args:
        base_dir (Path): 検索を行う基準ディレクトリ
        app_id (str): 探索するディレクトリ名のアプリID

    Returns:
        Path: 見つかったディレクトリのパス、見つからない場合はNone
    """
    return next((d for d in base_dir.iterdir() if d.is_dir() and d.name.startswith(f"{app_id}_")), None)

# ACLをExcelに変換
def generate_acl_excel(config, logger, app_id=None):
    """
    kintone_get_appjson の aclJson_to_excel.py を使用してACL情報をExcelに変換する
    
    Args:
        config (dict): 設定情報
        logger (Logger): ロガーオブジェクト
        app_id (int, optional): アプリID
    
    Returns:
        str: 生成されたExcelファイルのパス、または失敗した場合はFalse
    """
    logger.info("ACL情報のExcel変換を開始します")
    
    script_path = APPJSON_DIR / "aclJson_to_excel.py"
    
    if not script_path.exists():
        logger.error(f"スクリプトファイルが見つかりません: {script_path}")
        return False
    
    # app_tokensからアプリIDとAPIトークンを取得
    app_tokens = config.get('app_tokens', {})
    
    if app_id:
        # 特定のアプリIDが指定された場合
        if str(app_id) not in app_tokens:
            logger.error(f"アプリID {app_id} のAPIトークンが設定されていません")
            return False
            
        # [app_id]_ で始まるディレクトリを探す
        output_dir = find_existing_directory(OUTPUT_DIR, str(app_id))
        
        if not output_dir:
            logger.error(f"アプリID {app_id} に対応するディレクトリが見つかりません")
            return False
        
        output_file = output_dir / f"{app_id}_acl_report.xlsx"
        
        cmd = [
            sys.executable,
            str(script_path),
            str(app_id),
            "--output", str(output_file)
        ]
        
        try:
            logger.info(f"実行コマンド: python {script_path} {app_id} --output {output_file}")
            result = subprocess.run(cmd, check=True, capture_output=True, text=True)
            logger.info(f"アプリID {app_id} のACL情報を {output_file} に出力しました")
            logger.debug(f"出力: {result.stdout}")
            return str(output_file)
        except subprocess.CalledProcessError as e:
            logger.error(f"ACL情報のExcel変換中にエラーが発生しました: {e}")
            logger.error(f"標準出力: {e.stdout}")
            logger.error(f"標準エラー: {e.stderr}")
            return False
    else:
        # 全てのアプリを処理
        success = True
        generated_files = []
        
        for app_id in app_tokens.keys():
            # [app_id]_ で始まるディレクトリを探す
            output_dir = find_existing_directory(OUTPUT_DIR, str(app_id))
            
            if not output_dir:
                logger.error(f"アプリID {app_id} に対応するディレクトリが見つかりません")
                success = False
                continue
            
            output_file = output_dir / f"{app_id}_acl_report.xlsx"
            
            cmd = [
                sys.executable,
                str(script_path),
                str(app_id),
                "--output", str(output_file)
            ]
            
            try:
                logger.info(f"実行コマンド: python {script_path} {app_id} --output {output_file}")
                result = subprocess.run(cmd, check=True, capture_output=True, text=True)
                logger.info(f"アプリID {app_id} のACL情報を {output_file} に出力しました")
                logger.debug(f"出力: {result.stdout}")
                generated_files.append(str(output_file))
            except subprocess.CalledProcessError as e:
                logger.error(f"アプリID {app_id} のACL情報のExcel変換中にエラーが発生しました: {e}")
                logger.error(f"標準出力: {e.stdout}")
                logger.error(f"標準エラー: {e.stderr}")
                success = False
        
        if success and generated_files:
            return generated_files
        elif generated_files:
            return generated_files
        else:
            return False

# ディレクトリ操作関数
def prepare_directories():
    """
    ディレクトリの準備:
    1. PREVIOUS_OUTPUT_DIRを空にする
    2. OUTPUT_DIRの内容をPREVIOUS_OUTPUT_DIRに移動
    3. OUTPUT_DIRを作成
    """
    # 各ディレクトリが存在しない場合は作成
    for directory in [OUTPUT_DIR, PREVIOUS_OUTPUT_DIR, BACKUP_DIR]:
        directory.mkdir(exist_ok=True)
    
    # PREVIOUS_OUTPUT_DIRを空にする
    if PREVIOUS_OUTPUT_DIR.exists():
        for item in PREVIOUS_OUTPUT_DIR.iterdir():
            if item.is_file():
                item.unlink()
            elif item.is_dir():
                import shutil
                shutil.rmtree(item)
    
    # OUTPUT_DIRの内容をPREVIOUS_OUTPUT_DIRに移動
    if OUTPUT_DIR.exists():
        for item in OUTPUT_DIR.iterdir():
            if item.is_file():
                import shutil
                shutil.move(str(item), str(PREVIOUS_OUTPUT_DIR / item.name))
            elif item.is_dir():
                import shutil
                shutil.move(str(item), str(PREVIOUS_OUTPUT_DIR / item.name))
    
    # OUTPUT_DIRを作成（移動後に空になっている可能性があるため）
    OUTPUT_DIR.mkdir(exist_ok=True)

def backup_output():
    """
    OUTPUT_DIRの内容をBACKUP_DIRにバックアップする
    バックアップディレクトリ名: YYYYMMDD_HHMMSS
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_subdir = BACKUP_DIR / timestamp
    
    # バックアップディレクトリを作成
    backup_subdir.mkdir(exist_ok=True)
    
    # OUTPUT_DIRの内容をバックアップディレクトリにコピー
    if OUTPUT_DIR.exists():
        import shutil
        for item in OUTPUT_DIR.iterdir():
            if item.is_file():
                shutil.copy2(str(item), str(backup_subdir / item.name))
            elif item.is_dir():
                shutil.copytree(str(item), str(backup_subdir / item.name))
    
    return backup_subdir

def main():
    """メイン関数"""
    # コマンドライン引数の解析
    parser = argparse.ArgumentParser(description='Kintone関連ツールの統合実行スクリプト')
    subparsers = parser.add_subparsers(dest='command', help='実行するコマンド')
    
    # ユーザーグループ取得コマンド
    user_group_parser = subparsers.add_parser('users', help='ユーザーとグループ情報を取得（出力: kintone_users_groups_[日時].xlsx）')
    user_group_parser.add_argument('--format', choices=['excel', 'csv'], default='excel', help='出力形式')
    
    # アプリJSON取得コマンド
    app_json_parser = subparsers.add_parser('app', help='アプリのJSONデータを取得（出力: [アプリID]_app_settings.json, [アプリID]_form_layout.json など）')
    app_json_parser.add_argument('--id', type=int, help='取得するアプリID')
    
    # ACL Excel生成コマンド
    acl_excel_parser = subparsers.add_parser('acl', help='アプリのACL情報をExcelに変換（出力: acl_report_[アプリID]_[日時].xlsx）')
    acl_excel_parser.add_argument('--id', type=int, help='変換するアプリID')
    
    # グループ操作コマンド
    group_parser = subparsers.add_parser('group', help='グループ操作')
    group_subparsers = group_parser.add_subparsers(dest='action', help='実行するアクション')
    
    # グループ一覧
    group_subparsers.add_parser('list', help='グループ一覧を表示（コンソール出力）')
    
    # ユーザー検索
    search_parser = group_subparsers.add_parser('search', help='ユーザーを検索（コンソール出力）')
    search_parser.add_argument('keyword', help='検索キーワード')
    
    # ユーザーをグループに追加
    add_parser = group_subparsers.add_parser('add', help='ユーザーをグループに追加')
    add_parser.add_argument('user', help='ユーザーコード')
    add_parser.add_argument('group', help='グループ名またはコード')
    
    # ユーザーをグループから削除
    remove_parser = group_subparsers.add_parser('remove', help='ユーザーをグループから削除')
    remove_parser.add_argument('user', help='ユーザーコード')
    
    # 全機能実行コマンド
    subparsers.add_parser('all', help='すべての機能を順番に実行（複数の出力ファイルが生成されます）')
    
    # 出力ファイル一覧表示コマンド
    subparsers.add_parser('outputs', help='生成されるExcel/CSV/TSVファイルの一覧と概要を表示')
    
    # 環境ファイルオプション
    parser.add_argument('--env', type=str, help='.kintone.env ファイルのパス')
    
    # 引数がない場合はヘルプと出力ファイル情報を表示
    if len(sys.argv) == 1:
        parser.print_help()
        print("\n")
        display_output_info()
        sys.exit(0)
    
    args = parser.parse_args()
    
    # 出力ファイル一覧表示の場合
    if args.command == 'outputs':
        display_output_info()
        sys.exit(0)
    
    # ロギングの設定
    logger = setup_logging()
    logger.info("KintoneRunnerを起動しました")
    
    # ディレクトリの準備
    logger.info("ディレクトリの準備を開始します")
    prepare_directories()
    logger.info("ディレクトリの準備が完了しました")
    
    # 設定ファイルの読み込み
    env_file = Path(args.env) if args.env else ENV_FILE
    config = load_env_config(env_file)
    logger.info(f"設定ファイル {env_file} を読み込みました")
    
    # コマンドに応じて処理を実行
    if args.command == 'users':
        result = get_user_group_info(config, logger, args.format)
        if result:
            print(f"ユーザーとグループ情報を {result} に出力しました")
            
    elif args.command == 'app':
        result = get_app_json(config, logger, args.id)
        if result:
            print("アプリのJSONデータ取得が完了しました")
            
    elif args.command == 'acl':
        result = generate_acl_excel(config, logger, args.id)
        if result:
            if isinstance(result, list):
                print("以下のファイルにACL情報を出力しました:")
                for file in result:
                    print(f"- {file}")
            else:
                print(f"ACL情報を {result} に出力しました")
            
    elif args.command == 'group':
        if args.action == 'list':
            result = manage_groups(config, logger, 'list')
            if result:
                print(result)
                
        elif args.action == 'search':
            result = manage_groups(config, logger, 'search', {'keyword': args.keyword})
            if result:
                print(result)
                
        elif args.action == 'add':
            result = manage_groups(config, logger, 'add', {'user': args.user, 'group': args.group})
            if result:
                print(f"ユーザー {args.user} をグループ {args.group} に追加しました")
                
        elif args.action == 'remove':
            result = manage_groups(config, logger, 'remove', {'user': args.user})
            if result:
                print(f"ユーザー {args.user} をグループから削除しました")
                
    elif args.command == 'all':
        # すべての機能を順番に実行
        logger.info("すべての機能を順番に実行します")
        
        # 1. ユーザーとグループ情報の取得
        user_group_file = get_user_group_info(config, logger)
        if user_group_file:
            print(f"ユーザーとグループ情報を {user_group_file} に出力しました")
        
        # 2. アプリのJSONデータ取得
        app_result = get_app_json(config, logger)
        if app_result:
            print("アプリのJSONデータ取得が完了しました")
        
        # 3. ACL情報のExcel変換
        acl_result = generate_acl_excel(config, logger)
        if acl_result:
            if isinstance(acl_result, list):
                print("以下のファイルにACL情報を出力しました:")
                for file in acl_result:
                    print(f"- {file}")
            else:
                print(f"ACL情報を {acl_result} に出力しました")
        
        # 4. グループ一覧の表示
        group_result = manage_groups(config, logger, 'list')
        if group_result:
            print("グループ一覧:")
            print(group_result)
    
    # 処理完了後にバックアップを作成
    if args.command in ['users', 'app', 'acl', 'all']:
        backup_dir = backup_output()
        logger.info(f"出力ファイルを {backup_dir} にバックアップしました")
        print(f"出力ファイルを {backup_dir} にバックアップしました")
    
    logger.info("KintoneRunnerを終了します")

if __name__ == "__main__":
    main() 