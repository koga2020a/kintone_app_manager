#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Kintone関連ツールの統合実行スクリプト

このスクリプトは.kintone.envファイルから設定を読み込み、
kintone_get_user_group、kintone_get_appjson、kintone_group_cliの
機能を連携して実行するためのものです。

端末には進捗と結果だけを表示し、詳細は logs/kintone_runner_*.log に残します。
--verbose / -v を付けると INFO ログも端末に表示します。
"""

import os
import sys
import yaml
import argparse
import subprocess
import logging
import re
from pathlib import Path
from datetime import datetime

from console import DIM, arrow, c, fail, heading, note, ok, warn
from url_inventory import (
    build_inventory,
    export_inventory_to_excel,
    format_inventory_summary,
    normalize_domain_map,
)

# 定数定義
SCRIPT_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = SCRIPT_DIR / "output"
PREVIOUS_OUTPUT_DIR = SCRIPT_DIR / "previous_output"
BACKUP_DIR = SCRIPT_DIR / "backup"
ENV_FILE = SCRIPT_DIR / ".kintone.env"
ERROR_REPORT_FILE = SCRIPT_DIR / "error_report.txt"

# 各ディレクトリのパス
USER_GROUP_DIR = SCRIPT_DIR / "kintone_get_user_group"
APPJSON_DIR = SCRIPT_DIR / "kintone_get_appjson"
GROUP_CLI_DIR = SCRIPT_DIR / "kintone_group_cli"

# ユーザー・グループ一覧のマスタYAML（users コマンドが output/ 直下に生成する）
GROUP_USER_LIST_FILE = OUTPUT_DIR / "group_user_list.yaml"

# 子プロセスが失敗したときに端末へ出す出力の行数
CHILD_OUTPUT_TAIL_LINES = 30

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
            "name": "[アプリID]_acl_report.xlsx",
            "description": "アプリのACL情報（ユーザー名・グループ名を反映）※ output/[アプリID]_[アプリ名]/ の中",
            "command": "acl",
            "args": "--id [アプリID] (省略時は全アプリ対象)"
        },
        {
            "name": "kintone_app_settings_summary_[日時].xlsx",
            "description": "アプリの全体設定一覧表",
            "command": "summary",
            "args": "--output [ファイル名] (省略時は自動生成)"
        },
        {
            "name": "[アプリID]_notifications.xlsx",
            "description": "アプリの通知設定（一般・レコード・リマインダー）情報 ※ output/[アプリID]_[アプリ名]/ の中",
            "command": "notifications",
            "args": "--id [アプリID] (省略時は全アプリ対象)"
        },
        {
            "name": "[アプリID]_process_workflow.xlsx",
            "description": "アプリのプロセス管理（ワークフロー）情報 ※ output/[アプリID]_[アプリ名]/ の中",
            "command": "process_workflow",
            "args": "--id [アプリID] (省略時は全アプリ対象)"
        },
        {
            "name": "[アプリID]_layout_report.xlsx",
            "description": "フォームレイアウトの一覧表（取得時に自動生成）※ output/[アプリID]_[アプリ名]/ の中",
            "command": "app",
            "args": "--id [アプリID] (省略時は全アプリ対象)"
        },
        {
            "name": "search_hits_[日時].xlsx",
            "description": "取得済みデータの全文検索結果（検索語・アプリ毎のサマリとヒット一覧）",
            "command": "search",
            "args": "[検索語...] --id [アプリID] (--no-excel で出力しない)"
        },
        {
            "name": "url_inventory_[日時].xlsx",
            "description": "Webhook・JS の HTTP 呼び出し・外部参照 URL の横断一覧（ドメイン変更の as-is / to-be 確認用）",
            "command": "urls",
            "args": "--id [アプリID ...] --output [ファイル名] (省略時は取得済み全アプリ)"
        }
    ],
    "csv": [
        {
            "name": "kintone_users_groups_[日時].csv",
            "description": "ユーザーとグループの一覧情報（CSV形式・UTF-8 BOM付き）",
            "command": "users",
            "args": "--format csv"
        },
        {
            "name": "[アプリID]permission_target_user_names.csv",
            "description": "アプリに出現するユニークなユーザー名一覧（自動生成される補助ファイル）",
            "command": "acl",
            "args": "--id [アプリID] (省略時は全アプリ対象)"
        }
    ],
    "tsv": [
        {
            "name": "[アプリID]_layout_raw.tsv",
            "description": "フォームレイアウトの生データ（取得時に自動生成）※ output/[アプリID]_[アプリ名]/ の中",
            "command": "app",
            "args": "--id [アプリID] (省略時は全アプリ対象)"
        },
        {
            "name": "[アプリID]_layout_structured.tsv",
            "description": "フォームレイアウトを整形したデータ（取得時に自動生成）※ output/[アプリID]_[アプリ名]/ の中",
            "command": "app",
            "args": "--id [アプリID] (省略時は全アプリ対象)"
        }
    ]
}


# ---------------------------------------------------------------------------
# ログ・表示のヘルパー
# ---------------------------------------------------------------------------

def setup_logging(verbose=False):
    """
    ロギングの設定

    - ファイル (logs/kintone_runner_[日時].log): INFO 以上・従来フォーマット
    - 端末 (stderr): 既定は WARNING 以上、--verbose 指定時は INFO 以上

    Returns:
        tuple: (logger, log_file のパス)
    """
    log_dir = SCRIPT_DIR / "logs"
    log_dir.mkdir(exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = log_dir / f"kintone_runner_{timestamp}.log"

    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(
        logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    )

    stream_handler = logging.StreamHandler(sys.stderr)
    stream_handler.setLevel(logging.INFO if verbose else logging.WARNING)
    stream_handler.setFormatter(logging.Formatter('%(levelname)s: %(message)s'))

    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)
    for handler in list(root_logger.handlers):
        root_logger.removeHandler(handler)
    root_logger.addHandler(file_handler)
    root_logger.addHandler(stream_handler)

    return logging.getLogger("kintone_runner"), log_file


def rel_path(path):
    """スクリプトのあるディレクトリからの相対パスを返す（外にある場合はそのまま）"""
    try:
        return str(Path(path).resolve().relative_to(SCRIPT_DIR))
    except (ValueError, OSError):
        return str(path)


def mask_secret(value):
    """APIトークン・パスワードを UdLj***95eE の形にマスクする"""
    if value is None:
        return ""
    text = str(value)
    if len(text) <= 8:
        return "***"
    return f"{text[:4]}***{text[-4:]}"


def collect_secrets(config):
    """設定に含まれる秘密情報（パスワード・APIトークン）の実値リストを返す"""
    secrets = []
    if not isinstance(config, dict):
        return secrets
    password = config.get('password')
    if password:
        secrets.append(str(password))
    tokens = config.get('app_tokens') or {}
    if isinstance(tokens, dict):
        for token in tokens.values():
            if token:
                secrets.append(str(token))
    # 長いものから置換しないと部分一致で取りこぼす
    return sorted(set(secrets), key=len, reverse=True)


def mask_text(text, secrets=None):
    """文字列中の秘密情報をマスクした文字列を返す"""
    if not text:
        return text
    masked = str(text)
    for secret in (secrets or []):
        if secret:
            masked = masked.replace(secret, mask_secret(secret))
    return masked


def mask_command(command, secrets=None):
    """コマンド（list または str）を1行の文字列にし、秘密情報をマスクする"""
    if isinstance(command, (list, tuple)):
        command = ' '.join(str(part) for part in command)
    return mask_text(command, secrets)


def print_child_tail(stdout, stderr, limit=CHILD_OUTPUT_TAIL_LINES):
    """子プロセス出力の末尾だけを端末に薄く表示する（全文はログにある）"""
    lines = []
    for text in (stdout, stderr):
        if text and text.strip():
            lines.extend(text.rstrip().splitlines())
    if not lines:
        return
    if len(lines) > limit:
        skipped = len(lines) - limit
        print(c(f"    ...(前の {skipped} 行は省略。全文はログを参照)", DIM))
        lines = lines[-limit:]
    for line in lines:
        print(c(f"    {line}", DIM))


def run_child(cmd, logger, context, secrets=None):
    """
    子プロセスを実行する。

    出力は全文をログに残し、失敗時のみ端末に末尾を表示する。
    終了コード 3（APIトークン権限不足）のときは、子スクリプトの案内を
    そのまま端末に見せる。

    Returns:
        tuple: (成功したか(bool), 終了コード(int), stdout(str), stderr(str))
    """
    logger.info(f"実行コマンド: {mask_command(cmd, secrets)}")
    try:
        result = subprocess.run(cmd, check=True, capture_output=True, text=True, cwd=str(SCRIPT_DIR))
        if result.stdout:
            logger.info(f"標準出力:\n{mask_text(result.stdout, secrets)}")
        if result.stderr:
            logger.info(f"標準エラー:\n{mask_text(result.stderr, secrets)}")
        return True, 0, result.stdout or "", result.stderr or ""
    except subprocess.CalledProcessError as e:
        stdout = e.stdout or ""
        stderr = e.stderr or ""
        logger.error(f"{context} でエラーが発生しました（終了コード {e.returncode}）")
        logger.error(f"標準出力:\n{mask_text(stdout, secrets)}")
        logger.error(f"標準エラー:\n{mask_text(stderr, secrets)}")
        log_error_to_file(
            logger, e, command=cmd, stdout=stdout, stderr=stderr,
            context=context, secrets=secrets
        )
        if e.returncode == 3:
            # APIトークンの権限不足。子スクリプトの案内をそのまま見せる
            if stdout.strip():
                print(stdout.rstrip())
            if stderr.strip():
                print_child_tail(None, stderr)
        else:
            print_child_tail(stdout, stderr)
        return False, e.returncode, stdout, stderr
    except Exception as e:
        logger.error(f"{context} で予期しないエラーが発生しました: {e}")
        log_error_to_file(logger, e, command=cmd, context=context, secrets=secrets)
        return False, 1, "", str(e)


def print_result_summary(log_file, failures):
    """最後にまとめの1行とログの場所を出す"""
    print()
    if failures:
        fail(f"{failures} 件失敗")
    else:
        ok("すべて完了")
    if log_file:
        note(f"ログ: {rel_path(log_file)}")


# ---------------------------------------------------------------------------
# 設定ファイル
# ---------------------------------------------------------------------------

def load_env_config(env_file=None):
    """
    .kintone.env ファイルを読み込み、設定情報を返す
    """
    if env_file is None:
        env_file = ENV_FILE

    if not env_file.exists():
        fail(f"設定ファイル {env_file} が見つかりません。")
        sys.exit(1)

    try:
        with open(env_file, 'r', encoding='utf-8') as f:
            content = f.read()
            config = yaml.safe_load(content)

        # 必須項目をチェック
        required_keys = ['subdomain', 'username', 'password']
        missing_keys = [key for key in required_keys if key not in config]

        if missing_keys:
            fail(f"設定ファイルに以下の必須項目がありません: {', '.join(missing_keys)}")
            sys.exit(1)

        # app_tokens が辞書形式でない場合の処理
        if 'app_tokens' in config and config['app_tokens'] is None:
            config['app_tokens'] = {}

        return config
    except SystemExit:
        raise
    except Exception as e:
        fail(f"設定ファイルの読み込み中にエラーが発生しました: {e}")
        sys.exit(1)


def resolve_app_token(config, app_id):
    """app_tokens から指定アプリIDのトークンを取り出す（文字列キー・数値キーの両対応）"""
    tokens = config.get('app_tokens') or {}
    if not isinstance(tokens, dict):
        return None
    key = str(app_id)
    if key in tokens:
        return tokens[key]
    if key.isdigit() and int(key) in tokens:
        return tokens[int(key)]
    return None


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

    print("※ users / app / acl / summary / notifications / process_workflow は 'all' コマンドでも一括生成できます。")
    print("※ 出力先ディレクトリ: ./output/")


# エラー情報をファイルに記録する関数
def log_error_to_file(logger, error, command=None, stdout=None, stderr=None,
                      context=None, secrets=None):
    """
    エラー情報をerror_report.txtファイルに追記する

    Args:
        logger (Logger): ロガーオブジェクト
        error (Exception): 発生した例外
        command (str|list, optional): 実行されたコマンド
        stdout (str, optional): 標準出力の内容
        stderr (str, optional): 標準エラー出力の内容
        context (str, optional): エラーが発生した文脈（どの処理中か）
        secrets (list, optional): マスクすべき秘密情報の実値（パスワード・APIトークン）
    """
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    try:
        with open(ERROR_REPORT_FILE, 'a', encoding='utf-8') as f:
            f.write(f"===== エラーレポート: {timestamp} =====\n")

            if context:
                f.write(f"処理内容: {context}\n")

            if command:
                # パスワード・APIトークンをマスク
                f.write(f"実行コマンド: {mask_command(command, secrets)}\n")

            f.write(f"エラータイプ: {type(error).__name__}\n")
            f.write(f"エラーメッセージ: {mask_text(str(error), secrets)}\n")

            # トレースバック情報を追加
            import traceback
            tb_str = traceback.format_exc()
            f.write(f"\n--- トレースバック ---\n{mask_text(tb_str, secrets)}\n")

            if stdout:
                f.write(f"\n--- 標準出力 ---\n{mask_text(stdout, secrets)}\n")

            if stderr:
                f.write(f"\n--- 標準エラー出力 ---\n{mask_text(stderr, secrets)}\n")

            f.write("\n\n")

        logger.info(f"エラー情報を {ERROR_REPORT_FILE} に記録しました")
    except Exception as e:
        logger.error(f"エラー情報の記録中にエラーが発生しました: {e}")


# ---------------------------------------------------------------------------
# ユーザーとグループ情報の取得
# ---------------------------------------------------------------------------

def convert_users_excel_to_csv(xlsx_path, logger):
    """
    ユーザー一覧の xlsx を CSV に変換する。

    シートが1つなら kintone_users_groups_[日時].csv、
    複数なら kintone_users_groups_[日時]_[シート名].csv として保存する。

    Returns:
        list: 生成した CSV ファイルのパス（文字列）のリスト
    """
    import pandas as pd

    xlsx_path = Path(xlsx_path)
    sheets = pd.read_excel(xlsx_path, sheet_name=None)
    generated = []

    if len(sheets) == 1:
        df = next(iter(sheets.values()))
        csv_path = xlsx_path.with_suffix('.csv')
        df.to_csv(csv_path, index=False, encoding='utf-8-sig')
        generated.append(str(csv_path))
    else:
        for sheet_name, df in sheets.items():
            safe_name = re.sub(r'[\\/:*?"<>|\s]+', '_', str(sheet_name)).strip('_')
            csv_path = xlsx_path.parent / f"{xlsx_path.stem}_{safe_name}.csv"
            df.to_csv(csv_path, index=False, encoding='utf-8-sig')
            generated.append(str(csv_path))

    logger.info(f"CSVに変換しました: {', '.join(generated)}")
    return generated


def get_user_group_info(config, logger, output_format="excel"):
    """
    kintone_get_user_group の機能を呼び出してユーザーとグループ情報を取得

    Returns:
        tuple: (生成したファイルのパスのリスト, 失敗件数)
    """
    logger.info("ユーザーとグループ情報の取得を開始します")

    script_path = USER_GROUP_DIR / "get_user_group.py"

    if not script_path.exists():
        logger.error(f"スクリプトファイルが見つかりません: {script_path}")
        fail(f"スクリプトファイルが見つかりません: {rel_path(script_path)}")
        return [], 1

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
        "--output", str(output_file),
        "--yaml-dir", str(OUTPUT_DIR)
    ]

    print("  ユーザーとグループ情報を取得しています ...")
    success, _, _, _ = run_child(
        cmd, logger,
        context="ユーザーとグループ情報の取得",
        secrets=collect_secrets(config),
    )
    if not success:
        fail("ユーザーとグループ情報の取得に失敗しました")
        return [], 1

    generated = [str(output_file)]
    ok("ユーザーとグループ情報の取得が完了しました")
    arrow(rel_path(output_file))

    if output_format == "csv":
        try:
            csv_files = convert_users_excel_to_csv(output_file, logger)
            for csv_file in csv_files:
                arrow(rel_path(csv_file))
            generated = csv_files + generated
        except Exception as e:
            logger.error(f"CSVへの変換中にエラーが発生しました: {e}")
            log_error_to_file(
                logger, e, context="ユーザー一覧のCSV変換",
                secrets=collect_secrets(config),
            )
            fail(f"CSVへの変換に失敗しました: {e}")
            return generated, 1

    return generated, 0


# ---------------------------------------------------------------------------
# アプリJSONの取得
# ---------------------------------------------------------------------------

def get_app_json(config, logger, app_id=None):
    """
    アプリ設定の取得スクリプトを順に実行するラッパー関数

    download2yaml_excel.py（実際の取得）→ process_workflow_to_excel.py の順に実行する。
    最初のスクリプトがAPIトークン権限不足で失敗した場合は "forbidden" を伝播する。

    Returns:
        tuple: (状態("ok" または "forbidden"), 失敗件数)
    """
    scripts_to_run = [
        "download2yaml_excel.py",
        "process_workflow_to_excel.py",
    ]

    total_failed = 0
    for script in scripts_to_run:
        logger.info(f"==== スクリプト [{script}] の実行開始 ====")
        status, failed = get_app_json_do(
            config, logger, app_id=app_id, script_filename=script
        )
        if status == "forbidden":
            logger.error(f"スクリプト [{script}] の実行に失敗しました（APIトークン権限不足）")
            return "forbidden", failed + total_failed
        total_failed += failed
        if failed:
            logger.error(f"スクリプト [{script}] の実行に失敗しました")
            return "ok", total_failed
        logger.info(f"==== スクリプト [{script}] の実行完了 ====")

    return "ok", total_failed


def get_app_json_do(config, logger, app_id=None, script_filename="download2yaml_excel.py"):
    """
    指定したスクリプトを使ってアプリのJSONデータを取得／処理します。

    Returns:
        tuple: (状態("ok" または "forbidden"), 失敗件数)
    """
    logger.info(f"スクリプト [{script_filename}] を使ってアプリのJSONデータ取得を開始します")

    script_path = APPJSON_DIR / script_filename
    if not script_path.exists():
        logger.error(f"スクリプトファイルが見つかりません: {script_path}")
        fail(f"スクリプトファイルが見つかりません: {rel_path(script_path)}")
        return "ok", 1

    # 出力ディレクトリを準備
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # 設定から app_tokens を取得
    app_tokens = config.get('app_tokens', {})
    logger.info(
        "app_tokens: "
        + ", ".join(f"{k}: {mask_secret(v)}" for k, v in app_tokens.items())
    )

    # 処理対象リストを作成
    if app_id is not None:
        token = resolve_app_token(config, app_id)
        if token is None:
            logger.error(f"アプリID {app_id} の API トークンが設定されていません")
            fail(f"アプリID {app_id} の APIトークンが設定されていません")
            return "ok", 1
        items = [(str(app_id), token)]
    else:
        items = [(str(k), v) for k, v in app_tokens.items()]

    if not items:
        warn("対象のアプリがありません（.kintone.env の app_tokens が空です）")
        return "ok", 0

    secrets = collect_secrets(config)
    total = len(items)
    failed = 0

    for i, (aid, api_token) in enumerate(items, start=1):
        logger.info(f"アプリID {aid} の処理を開始します")
        print(f"  [{i}/{total}] アプリID {aid} ... ({script_filename})")
        cmd = [
            sys.executable,
            str(script_path),
            aid,
            api_token,
            config["subdomain"],
            config["username"],
            config["password"],
        ]

        success, returncode, _, _ = run_child(
            cmd, logger,
            context=f"アプリID {aid} のJSONデータ取得",
            secrets=secrets,
        )
        if success:
            ok(f"アプリID {aid} 完了")
            app_dir = find_existing_directory(OUTPUT_DIR, aid)
            if app_dir:
                arrow(rel_path(app_dir))
        else:
            fail(f"アプリID {aid} 失敗")
            if returncode == 3:
                return "forbidden", failed + 1
            failed += 1

    return "ok", failed


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
    if not base_dir.exists():
        return None
    return next((d for d in base_dir.iterdir() if d.is_dir() and d.name.startswith(f"{app_id}_")), None)


# ---------------------------------------------------------------------------
# グループ操作
# ---------------------------------------------------------------------------

def manage_groups(config, logger, action, params=None, env_file=None):
    """
    kintone_group_cli の機能を呼び出してグループを操作

    action: 'list', 'search', 'add', 'remove'
    params: アクションに応じたパラメータ
    env_file: group_cli.py に渡す設定ファイル（省略時は .kintone.env）

    Returns:
        tuple: (子プロセスの標準出力(str) または False, 失敗件数)
    """
    logger.info(f"グループ操作 '{action}' を開始します")

    script_path = GROUP_CLI_DIR / "group_cli.py"

    if not script_path.exists():
        logger.error(f"スクリプトファイルが見つかりません: {script_path}")
        fail(f"スクリプトファイルが見つかりません: {rel_path(script_path)}")
        return False, 1

    # 設定ファイルは .kintone.env をそのまま渡す（subdomain/username/password を含む）
    config_path = Path(env_file) if env_file else ENV_FILE
    if not config_path.exists():
        logger.error(f"設定ファイルが見つかりません: {config_path}")
        fail(f"設定ファイルが見つかりません: {rel_path(config_path)}")
        return False, 1

    cmd = [sys.executable, str(script_path), "--config", str(config_path.resolve())]

    # アクションに応じてコマンドラインを構築
    if action == 'list':
        cmd.append('list')
    elif action == 'search':
        if not params or 'keyword' not in params:
            logger.error("検索にはキーワードが必要です")
            fail("検索にはキーワードが必要です")
            return False, 1
        cmd.append('--search')
        cmd.append(params['keyword'])
    elif action == 'add':
        if not params or 'user' not in params or 'group' not in params:
            logger.error("ユーザー追加にはユーザーコードとグループ名/コードが必要です")
            fail("ユーザー追加にはユーザーコードとグループ名/コードが必要です")
            return False, 1
        cmd.extend(['set', params['user'], params['group']])
    elif action == 'remove':
        if not params or 'user' not in params:
            logger.error("ユーザー削除にはユーザーコードが必要です")
            fail("ユーザー削除にはユーザーコードが必要です")
            return False, 1
        cmd.extend(['set', params['user']])
    else:
        logger.error(f"不明なアクション: {action}")
        fail(f"不明なアクション: {action}")
        return False, 1

    success, _, stdout, _ = run_child(
        cmd, logger,
        context=f"グループ操作 '{action}'",
        secrets=collect_secrets(config),
    )

    if not success:
        fail(f"グループ操作 '{action}' に失敗しました")
        return False, 1

    return stdout, 0


# ---------------------------------------------------------------------------
# 取得済みデータのExcel変換
# ---------------------------------------------------------------------------

_group_master_warned = False


def warn_if_no_group_master():
    """ユーザー・グループ一覧のマスタが無ければ1回だけ警告する（処理は続行）"""
    global _group_master_warned
    if _group_master_warned or GROUP_USER_LIST_FILE.exists():
        return
    _group_master_warned = True
    warn(
        "ユーザー・グループ一覧が未取得のため、グループ名・ユーザー名はコードのまま"
        "出力します。先に users を実行してください"
    )


def generate_acl_excel(config, logger, app_id=None):
    """ACL情報をExcelに変換する"""
    return generate_excel_proc(
        config, logger, app_id=app_id,
        script_filename="aclJson_to_excel.py",
        excel_filename="acl_report.xlsx",
        label="ACL情報",
        extra_args=["--group-master", str(GROUP_USER_LIST_FILE)],
    )


def generate_notifications_excel(config, logger, app_id=None):
    """通知設定をExcelに変換する"""
    return generate_excel_proc(
        config, logger, app_id=app_id,
        script_filename="notifications_to_excel.py",
        excel_filename="notifications.xlsx",
        label="通知設定",
    )


def generate_process_workflow_excel(config, logger, app_id=None):
    """プロセスワークフローをExcelに変換する"""
    return generate_excel_proc(
        config, logger, app_id=app_id,
        script_filename="process_workflow_to_excel.py",
        excel_filename="process_workflow.xlsx",
        label="プロセスワークフロー",
    )


def generate_excel_proc(config, logger, app_id=None,
                        script_filename="notifications_to_excel.py",
                        excel_filename="notifications.xlsx",
                        label="設定",
                        extra_args=None):
    """
    kintone_get_appjson の各スクリプトを使って、取得済みデータをExcelに変換する

    Args:
        config (dict): 設定情報
        logger (Logger): ロガーオブジェクト
        app_id (int, optional): アプリID（省略時は app_tokens の全アプリ）
        script_filename (str): 実行するスクリプト名
        excel_filename (str): 出力ファイル名の接尾辞
        label (str): 端末表示用のラベル
        extra_args (list, optional): 子スクリプトに追加で渡す引数

    Returns:
        tuple: (生成したファイルのパスのリスト, 失敗件数)
    """
    logger.info(f"{label}のExcel変換を開始します（{script_filename}）")

    script_path = APPJSON_DIR / script_filename

    if not script_path.exists():
        logger.error(f"スクリプトファイルが見つかりません: {script_path}")
        fail(f"スクリプトファイルが見つかりません: {rel_path(script_path)}")
        return [], 1

    app_tokens = config.get('app_tokens', {})
    logger.info(
        "app_tokens: "
        + ", ".join(f"{k}: {mask_secret(v)}" for k, v in app_tokens.items())
    )

    # 処理対象のアプリIDを決める
    if app_id is not None:
        if resolve_app_token(config, app_id) is None:
            logger.error(f"アプリID {app_id} のAPIトークンが設定されていません")
            fail(f"アプリID {app_id} の APIトークンが設定されていません")
            return [], 1
        targets = [str(app_id)]
    else:
        targets = [str(k) for k in app_tokens.keys()]

    if not targets:
        warn("対象のアプリがありません（.kintone.env の app_tokens が空です）")
        return [], 0

    secrets = collect_secrets(config)
    total = len(targets)
    generated = []
    failed = 0

    for i, aid in enumerate(targets, start=1):
        print(f"  [{i}/{total}] アプリID {aid} ...")

        # [app_id]_ で始まるディレクトリを探す
        output_dir = find_existing_directory(OUTPUT_DIR, aid)
        if not output_dir:
            logger.error(f"アプリID {aid} に対応するディレクトリが見つかりません")
            fail(f"アプリID {aid} の取得済みフォルダがありません。先に app を実行してください")
            failed += 1
            continue

        output_file = output_dir / f"{aid}_{excel_filename}"
        cmd = [
            sys.executable,
            str(script_path),
            aid,
            "--output", str(output_file),
        ]
        if extra_args:
            cmd.extend(extra_args)

        success, _, _, _ = run_child(
            cmd, logger,
            context=f"アプリID {aid} の{label}のExcel変換",
            secrets=secrets,
        )
        if success and output_file.exists():
            logger.info(f"アプリID {aid} の{label}を {output_file} に出力しました")
            ok(f"アプリID {aid} 完了")
            arrow(rel_path(output_file))
            generated.append(str(output_file))
        elif success:
            logger.error(f"アプリID {aid} の出力ファイルが作成されませんでした: {output_file}")
            fail(f"アプリID {aid} 失敗（出力ファイルが作成されませんでした）")
            failed += 1
        else:
            fail(f"アプリID {aid} 失敗")
            failed += 1

    return generated, failed


def generate_summary_excel(config, logger, output=None):
    """
    app_settings_summary.py を使ってアプリ設定一覧表を生成する

    Returns:
        tuple: (生成したファイルのパスのリスト, 失敗件数)
    """
    logger.info("アプリ設定一覧表の生成を開始します")
    script_path = SCRIPT_DIR / "app_settings_summary.py"

    if not script_path.exists():
        logger.error(f"スクリプトファイルが見つかりません: {script_path}")
        fail(f"スクリプトファイル {rel_path(script_path)} が見つかりません")
        return [], 1

    cmd = [sys.executable, str(script_path)]
    if output:
        cmd.extend(["--output", output])

    print("  アプリ設定一覧表を作成しています ...")
    success, _, stdout, _ = run_child(
        cmd, logger,
        context="アプリ設定一覧表の生成",
        secrets=collect_secrets(config),
    )
    if not success:
        fail("アプリ設定一覧表の生成に失敗しました")
        return [], 1

    logger.info("アプリ設定一覧表の生成が完了しました")
    ok("アプリ設定一覧表の生成が完了しました")

    generated = []
    match = re.search(r'アプリ設定一覧表を (.+?) に出力しました', stdout or "")
    if match:
        generated.append(match.group(1).strip())
        arrow(rel_path(match.group(1).strip()))
    elif stdout and stdout.strip():
        for line in stdout.strip().splitlines():
            print(c(f"    {line}", DIM))

    return generated, 0


# ---------------------------------------------------------------------------
# 取得済みデータの検索 / Webhook 一覧 / URL 一覧
# ---------------------------------------------------------------------------

def command_search(config, logger, keywords, app_id=None, no_excel=False, verbose=False):
    """
    取得済みデータ（YAML/JSON/JS など）を全文検索する

    Returns:
        int: 失敗件数
    """
    from inspect_downloaded import (
        export_search_hits_to_excel,
        find_app_output_dir,
        format_search_hits,
        list_app_output_dirs,
        search_all_apps,
    )

    heading("取得済みデータの全文検索")

    words = [str(k).strip() for k in (keywords or []) if str(k).strip()]
    if not words:
        saved = config.get('search_keywords') or []
        if isinstance(saved, str):
            saved = [saved]
        words = [str(k).strip() for k in saved if str(k).strip()]
        if words:
            note("検索語は .kintone.env の search_keywords を使います")
    if not words:
        logger.error("検索語がありません")
        fail("検索語がありません。引数で指定するか .kintone.env の search_keywords に登録してください")
        sys.exit(2)

    if app_id is not None:
        app_dir = find_app_output_dir(OUTPUT_DIR, str(app_id))
        targets = [(str(app_id), app_dir)] if app_dir else []
    else:
        targets = list_app_output_dirs(OUTPUT_DIR)

    if not targets:
        logger.error("取得済みフォルダがありません")
        fail("取得済みフォルダがありません。先に app を実行してください")
        sys.exit(1)

    note(f"対象: {len(targets)} アプリ [{', '.join(aid for aid, _ in targets)}]")
    logger.info(f"検索対象: {[str(path) for _, path in targets]} / 検索語: {words}")

    hits = search_all_apps(targets, words)
    logger.info(f"検索結果: {len(hits)} 行")

    print()
    for line in format_search_hits(words, hits, multi_app=len(targets) > 1, detail=verbose):
        print(line)

    if not no_excel:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_path = OUTPUT_DIR / f"search_hits_{timestamp}.xlsx"
        try:
            export_search_hits_to_excel(words, hits, excel_path)
            print()
            ok("検索結果をExcelに出力しました")
            arrow(rel_path(excel_path))
        except Exception as e:
            logger.error(f"検索結果のExcel出力中にエラーが発生しました: {e}")
            log_error_to_file(
                logger, e, context="検索結果のExcel出力",
                secrets=collect_secrets(config),
            )
            fail(f"検索結果のExcel出力に失敗しました: {e}")
            return 1

    return 0


def command_webhooks(config, logger, app_id, verbose=False):
    """
    アプリの Webhook 設定を一覧表示する
    （保存済みファイルがあればそれを、無ければ内部APIで取得して保存する）

    Returns:
        int: 失敗件数
    """
    from inspect_downloaded import (
        extract_webhook_rows,
        fetch_and_save_webhooks,
        find_app_output_dir,
        find_webhook_file,
        format_search_hits,
        format_webhook_rows,
        is_stale_webhook_file,
        load_json_or_yaml,
        search_downloaded,
    )

    heading("Webhook設定一覧")

    app_id = str(app_id)
    app_dir = find_app_output_dir(OUTPUT_DIR, app_id)
    if not app_dir:
        logger.error(f"アプリID {app_id} の取得済みフォルダが見つかりません")
        fail(f"アプリID {app_id} の取得済みフォルダがありません。先に app を実行してください")
        return 1

    note(f"参照: {rel_path(app_dir)}")

    data = None
    webhook_file = find_webhook_file(app_dir, app_id)
    if webhook_file:
        try:
            data = load_json_or_yaml(webhook_file)
            note(f"ファイル: {webhook_file.relative_to(app_dir)}")
        except Exception as e:
            logger.warning(f"Webhookファイルの読み込みに失敗しました: {e}")
            warn(f"読み込みエラー: {e}")
            data = None
        if data is not None and is_stale_webhook_file(data):
            warn("保存済みファイルは旧方式（管理画面HTMLの走査）で取得したもののため再取得します。")
            data = None

    if data is None:
        note(f"管理画面と同じ内部API（/k/api/dev/app/{app_id}/webhook/list.json）で取得します。")
        logger.info(f"アプリID {app_id} の Webhook 一覧を内部APIで取得します")
        saved, data, error = fetch_and_save_webhooks(
            app_dir,
            app_id,
            subdomain=config.get('subdomain') or "",
            api_token=resolve_app_token(config, app_id),
            username=config.get('username'),
            password=config.get('password'),
        )
        if saved and data is not None:
            ok(f"取得して保存しました: {saved.relative_to(app_dir)}")
        else:
            logger.error(f"Webhook一覧の取得に失敗しました: {error}")
            warn("Webhook一覧APIからは取得できませんでした。")
            if error:
                note(f"({error})")
            note("取得済みファイル内の webhook という文字列を代わりに表示します。")
            hits = search_downloaded(app_dir, ["webhook", "Webhook", "WEBHOOK"])
            print()
            if hits:
                for line in format_search_hits(["webhook"], hits, detail=verbose):
                    print(line)
            else:
                print("  取得済みデータ内にも webhook の記載はありません（0件）。")
            return 1

    rows = extract_webhook_rows(data)
    logger.info(f"Webhook {len(rows)} 件")
    print()
    for line in format_webhook_rows(rows):
        print(line)

    return 0


def command_urls(config, logger, app_ids=None, output=None, verbose=False):
    """
    取得済みデータから Webhook・JS の HTTP 呼び出し・外部参照 URL を集めて Excel に出力する
    （.kintone.env の domain_map があれば「変更後 URL」と「状態」も付ける）

    Returns:
        int: 失敗件数
    """
    from inspect_downloaded import list_app_output_dirs

    heading("URL 一覧（Webhook / JS 呼び出し / 外部参照）")

    targets = list_app_output_dirs(OUTPUT_DIR)
    if app_ids:
        wanted = [str(app_id) for app_id in app_ids]
        available = {app_id for app_id, _ in targets}
        for app_id in wanted:
            if app_id not in available:
                logger.warning(f"アプリID {app_id} の取得済みフォルダがありません")
                warn(f"アプリID {app_id} は取得済みフォルダが無いため対象から外します")
        targets = [(app_id, path) for app_id, path in targets if app_id in set(wanted)]

    if not targets:
        logger.error("取得済みフォルダがありません")
        fail("取得済みフォルダがありません。先に app を実行してください")
        sys.exit(1)

    note(f"対象: {len(targets)} アプリ [{', '.join(app_id for app_id, _ in targets)}]")
    logger.info(f"URL 一覧の対象: {[str(path) for _, path in targets]}")

    domain_map = normalize_domain_map(config.get("domain_map"))
    if domain_map:
        note(f"ドメイン対応表: {len(domain_map)} 件")
        logger.info(f"domain_map: {domain_map}")
    else:
        note("domain_map が未設定のため、変更後 URL と状態は出力しません"
             "（.kintone.env の domain_map に 変更前ホスト: 変更後ホスト を登録）")

    inventory = build_inventory(targets, domain_map)
    logger.info(f"URL 一覧: {len(getattr(inventory, 'rows', []) or [])} 行")

    print()
    for line in format_inventory_summary(inventory):
        print(line)

    missing = list(getattr(inventory, "missing_webhooks", None) or [])
    if missing:
        # 端末への注意は format_inventory_summary が出すので、ここはログのみ
        logger.info(f"Webhook 未取得のアプリ: {missing}")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    if output:
        excel_path = Path(output)
    else:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_path = OUTPUT_DIR / f"url_inventory_{timestamp}.xlsx"

    try:
        saved = export_inventory_to_excel(inventory, excel_path)
        print()
        ok("URL 一覧をExcelに出力しました")
        arrow(rel_path(saved or excel_path))
    except Exception as e:
        logger.error(f"URL 一覧のExcel出力中にエラーが発生しました: {e}")
        log_error_to_file(
            logger, e, context="URL 一覧のExcel出力",
            secrets=collect_secrets(config),
        )
        fail(f"URL 一覧のExcel出力に失敗しました: {e}")
        return 1

    return 0


# ---------------------------------------------------------------------------
# ディレクトリ操作関数
# ---------------------------------------------------------------------------

def prepare_directories():
    """
    ディレクトリの準備:
    1. PREVIOUS_OUTPUT_DIRを空にする
    2. OUTPUT_DIRの内容をPREVIOUS_OUTPUT_DIRに移動
    3. OUTPUT_DIRを作成
    """
    import shutil
    import logging
    logger = logging.getLogger("kintone_runner")

    # Excelファイルが開かれているかどうかを確認するフラグ
    excel_files_open = False
    excel_files_list = []

    # 各ディレクトリが存在しない場合は作成
    for directory in [OUTPUT_DIR, PREVIOUS_OUTPUT_DIR, BACKUP_DIR]:
        directory.mkdir(exist_ok=True)

    # PREVIOUS_OUTPUT_DIRを空にする
    if PREVIOUS_OUTPUT_DIR.exists():
        for item in PREVIOUS_OUTPUT_DIR.iterdir():
            if item.is_file():
                try:
                    item.unlink()
                except PermissionError:
                    if item.name.startswith("~$"):
                        excel_files_open = True
                        excel_files_list.append(item.name[2:])  # "~$"を除いたファイル名
                        logger.warning(f"ファイル {item.name[2:]} はExcelで開かれているため削除できません。")
                    else:
                        logger.warning(f"ファイル {item.name} へのアクセスが拒否されました。")
            elif item.is_dir():
                try:
                    shutil.rmtree(item)
                except (PermissionError, OSError) as e:
                    logger.warning(f"ディレクトリ {item.name} の削除中にエラーが発生しました: {e}")

    # OUTPUT_DIRの内容をPREVIOUS_OUTPUT_DIRに移動
    if OUTPUT_DIR.exists():
        for item in OUTPUT_DIR.iterdir():
            try:
                if item.is_file():
                    # Excelの一時ファイルをチェック
                    if item.name.startswith("~$"):
                        excel_files_open = True
                        excel_files_list.append(item.name[2:])  # "~$"を除いたファイル名
                        logger.warning(f"Excelファイル {item.name[2:]} が開かれています。")
                        continue
                    shutil.move(str(item), str(PREVIOUS_OUTPUT_DIR / item.name))
                elif item.is_dir():
                    # ディレクトリ内にExcelの一時ファイルがないか確認
                    for file in item.glob("~$*"):
                        excel_files_open = True
                        excel_files_list.append(file.name[2:])  # "~$"を除いたファイル名
                        logger.warning(f"ディレクトリ {item.name} 内のExcelファイル {file.name[2:]} が開かれています。")

                    # Excelが開かれていない場合は通常通り移動
                    shutil.move(str(item), str(PREVIOUS_OUTPUT_DIR / item.name))
            except (PermissionError, OSError) as e:
                if "~$" in str(e):
                    excel_files_open = True
                    logger.warning(f"Excelファイルが開かれているため、ファイルを移動できませんでした。")
                else:
                    logger.warning(f"ファイルまたはディレクトリの移動中にエラーが発生しました: {e}")

    # Excelファイルが開かれている場合は例外を発生させる
    if excel_files_open:
        files_str = ", ".join(excel_files_list)
        error_msg = f"以下のExcelファイルが開かれているため処理を続行できません: {files_str}"
        logger.error(error_msg)
        raise PermissionError(error_msg)

    # OUTPUT_DIRを作成（移動後に空になっている可能性があるため）
    OUTPUT_DIR.mkdir(exist_ok=True)
    logger.info("ディレクトリの準備が完了しました。")

# 特定のアプリIDに関連するディレクトリのみを準備する関数
def prepare_app_directories(app_id):
    """
    特定のアプリID向けのディレクトリ準備:
    1. PREVIOUS_OUTPUT_DIRの指定アプリIDのディレクトリのみを削除
    2. OUTPUT_DIRの指定アプリIDのディレクトリをPREVIOUS_OUTPUT_DIRに移動
    3. OUTPUT_DIRを作成

    Args:
        app_id (int): 処理対象のアプリID
    """
    import shutil
    import logging
    logger = logging.getLogger("kintone_runner")

    # Excelファイルが開かれているかどうかを確認するフラグ
    excel_files_open = False
    excel_files_list = []

    # 各ディレクトリが存在しない場合は作成
    for directory in [OUTPUT_DIR, PREVIOUS_OUTPUT_DIR, BACKUP_DIR]:
        directory.mkdir(exist_ok=True)

    # PREVIOUS_OUTPUT_DIRの指定アプリIDのディレクトリのみを削除
    if PREVIOUS_OUTPUT_DIR.exists():
        app_dir = find_existing_directory(PREVIOUS_OUTPUT_DIR, str(app_id))
        if app_dir and app_dir.exists():
            try:
                for file in app_dir.glob("~$*"):
                    excel_files_open = True
                    excel_files_list.append(file.name[2:])
                    logger.warning(f"ディレクトリ {app_dir.name} 内のExcelファイル {file.name[2:]} が開かれています。")

                if not excel_files_open:
                    shutil.rmtree(app_dir)
                    logger.info(f"PREVIOUS_OUTPUT_DIRから {app_dir.name} を削除しました")
            except (PermissionError, OSError) as e:
                if "~$" in str(e):
                    excel_files_open = True
                    logger.warning(f"Excelファイルが開かれているため、ディレクトリを削除できませんでした。")
                else:
                    logger.warning(f"ディレクトリ {app_dir.name} の削除中にエラーが発生しました: {e}")

    # OUTPUT_DIRの指定アプリIDのディレクトリをPREVIOUS_OUTPUT_DIRに移動
    if OUTPUT_DIR.exists():
        app_dir = find_existing_directory(OUTPUT_DIR, str(app_id))
        if app_dir and app_dir.exists():
            try:
                # ディレクトリ内にExcelの一時ファイルがないか確認
                for file in app_dir.glob("~$*"):
                    excel_files_open = True
                    excel_files_list.append(file.name[2:])
                    logger.warning(f"ディレクトリ {app_dir.name} 内のExcelファイル {file.name[2:]} が開かれています。")

                # Excelが開かれていない場合は移動
                if not excel_files_open:
                    shutil.move(str(app_dir), str(PREVIOUS_OUTPUT_DIR / app_dir.name))
                    logger.info(f"OUTPUT_DIRから {app_dir.name} をPREVIOUS_OUTPUT_DIRに移動しました")
            except (PermissionError, OSError) as e:
                if "~$" in str(e):
                    excel_files_open = True
                    logger.warning(f"Excelファイルが開かれているため、ディレクトリを移動できませんでした。")
                else:
                    logger.warning(f"ディレクトリ {app_dir.name} の移動中にエラーが発生しました: {e}")

    # Excelファイルが開かれている場合は例外を発生させる
    if excel_files_open:
        files_str = ", ".join(excel_files_list)
        error_msg = f"以下のExcelファイルが開かれているため処理を続行できません: {files_str}"
        logger.error(error_msg)
        raise PermissionError(error_msg)

    # OUTPUT_DIRを作成（移動後に空になっている可能性があるため）
    OUTPUT_DIR.mkdir(exist_ok=True)
    logger.info(f"アプリID {app_id} のディレクトリ準備が完了しました。")

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

def remove_datetime_suffix(directory):
    """
    出力ディレクトリ内のファイル名とディレクトリ名から日時部分を除去する

    Args:
        directory (Path): 処理対象のディレクトリ
    """
    logger = logging.getLogger("kintone_runner")
    logger.info("ファイル名とディレクトリ名から日時部分を除去します")

    # 日時パターン（_YYYYMMDD_HHMMSS）を定義
    datetime_pattern = re.compile(r'_\d{8}_\d{6}')
    try:
        # ディレクトリ内のすべてのファイルとディレクトリを処理
        for item in directory.iterdir():
            original_name = item.name
            # 日時部分を除去
            new_name = datetime_pattern.sub('', original_name)

            if new_name != original_name:
                try:
                    new_path = item.parent / new_name
                    # 同名のファイルが存在する場合は上書き
                    if new_path.exists():
                        if new_path.is_file():
                            new_path.unlink()
                        else:
                            import shutil
                            shutil.rmtree(new_path)
                    item.rename(new_path)
                    logger.info(f"リネーム: {original_name} -> {new_name}")
                except Exception as e:
                    logger.error(f"リネーム中にエラーが発生しました ({original_name}): {e}")

        logger.info("ファイル名とディレクトリ名からの日時部分の除去が完了しました")
    except Exception as e:
        logger.error(f"ファイル名とディレクトリ名の処理中にエラーが発生しました: {e}")


# ---------------------------------------------------------------------------
# コマンドライン
# ---------------------------------------------------------------------------

def build_parser():
    """パーサを組み立てて (parser, group_parser) を返す"""
    # --env / --verbose は共通の親パーサにして、メイン・各サブコマンドの両方で受け付ける。
    # default=argparse.SUPPRESS にしておかないと、サブパーサの既定値が
    # メインパーサで指定された値を上書きしてしまう。
    common_parser = argparse.ArgumentParser(add_help=False)
    common_parser.add_argument(
        '--env', type=str, default=argparse.SUPPRESS,
        help='.kintone.env ファイルのパス'
    )
    common_parser.add_argument(
        '-v', '--verbose', action='store_true', default=argparse.SUPPRESS,
        help='INFOログも端末に表示する（検索は内訳も表示）'
    )

    parser = argparse.ArgumentParser(
        description='Kintone関連ツールの統合実行スクリプト',
        parents=[common_parser],
        epilog='対話メニュー: python menu.py',
    )
    subparsers = parser.add_subparsers(dest='command', help='実行するコマンド')

    # ユーザーグループ取得コマンド
    user_group_parser = subparsers.add_parser(
        'users', parents=[common_parser],
        help='ユーザーとグループ情報を取得（出力: kintone_users_groups_[日時].xlsx）')
    user_group_parser.add_argument('--format', choices=['excel', 'csv'], default='excel', help='出力形式')

    # アプリJSON取得コマンド
    app_json_parser = subparsers.add_parser(
        'app', parents=[common_parser],
        help='アプリのJSONデータを取得（出力: [アプリID]_app_settings.json, [アプリID]_form_layout.json など）')
    app_json_parser.add_argument('--id', type=int, help='取得するアプリID')

    # ACL Excel生成コマンド
    acl_excel_parser = subparsers.add_parser(
        'acl', parents=[common_parser],
        help='アプリのACL情報をExcelに変換（出力: [アプリID]_acl_report.xlsx）')
    acl_excel_parser.add_argument('--id', type=int, help='変換するアプリID')

    # アプリ設定一覧表生成コマンド
    summary_parser = subparsers.add_parser(
        'summary', parents=[common_parser],
        help='アプリの全体設定一覧表をExcelで出力（出力: kintone_app_settings_summary_[日時].xlsx）')
    summary_parser.add_argument('--output', type=str, help='出力ファイル名')

    # グループ操作コマンド
    group_parser = subparsers.add_parser('group', parents=[common_parser], help='グループ操作')
    group_subparsers = group_parser.add_subparsers(dest='action', help='実行するアクション')

    # グループ一覧
    group_subparsers.add_parser('list', parents=[common_parser], help='グループ一覧を表示（コンソール出力）')

    # ユーザー検索
    search_parser = group_subparsers.add_parser(
        'search', parents=[common_parser], help='ユーザーを検索（コンソール出力）')
    search_parser.add_argument('keyword', help='検索キーワード')

    # ユーザーをグループに追加
    add_parser = group_subparsers.add_parser(
        'add', parents=[common_parser], help='ユーザーをグループに追加')
    add_parser.add_argument('user', help='ユーザーコード')
    add_parser.add_argument('group', help='グループ名またはコード')

    # ユーザーをグループから削除
    remove_parser = group_subparsers.add_parser(
        'remove', parents=[common_parser], help='ユーザーをグループから削除')
    remove_parser.add_argument('user', help='ユーザーコード')

    # 通知設定Excel生成コマンド
    notifications_parser = subparsers.add_parser(
        'notifications', parents=[common_parser],
        help='アプリの通知設定をExcelに変換（出力: [アプリID]_notifications.xlsx）')
    notifications_parser.add_argument('--id', type=int, help='変換するアプリID')

    # プロセスワークフローExcel生成コマンド
    process_workflow_parser = subparsers.add_parser(
        'process_workflow', parents=[common_parser],
        help='アプリのプロセスワークフローをExcelに変換（出力: [アプリID]_process_workflow.xlsx）')
    process_workflow_parser.add_argument('--id', type=int, help='変換するアプリID')

    # 取得済みデータの全文検索コマンド
    search_downloaded_parser = subparsers.add_parser(
        'search', parents=[common_parser],
        help='取得済みデータ（YAML/JSON/JS）を全文検索（出力: search_hits_[日時].xlsx）')
    search_downloaded_parser.add_argument(
        'keywords', nargs='*',
        help='検索語（省略時は .kintone.env の search_keywords を使用）')
    search_downloaded_parser.add_argument('--id', type=int, help='検索するアプリID（省略時は取得済み全アプリ）')
    search_downloaded_parser.add_argument('--no-excel', action='store_true', help='Excelファイルを出力しない')

    # Webhook 一覧コマンド
    webhooks_parser = subparsers.add_parser(
        'webhooks', parents=[common_parser],
        help='アプリの Webhook 設定を一覧表示（取得済みが無ければ取得して保存）')
    webhooks_parser.add_argument('--id', type=int, required=True, help='表示するアプリID')

    # URL 一覧コマンド
    urls_parser = subparsers.add_parser(
        'urls', parents=[common_parser],
        help='Webhook・JS の HTTP 呼び出し・外部参照 URL の一覧を Excel 出力'
             '（ドメイン変更の as-is / to-be 確認用。出力: url_inventory_[日時].xlsx）')
    urls_parser.add_argument(
        '--id', type=int, nargs='+', help='対象とするアプリID（省略時は取得済み全アプリ）')
    urls_parser.add_argument('--output', type=str, help='出力ファイル名')

    # 全機能実行コマンド
    all_parser = subparsers.add_parser(
        'all', parents=[common_parser],
        help='すべての機能を順番に実行（複数の出力ファイルが生成されます）')
    all_parser.add_argument('--id', type=int, nargs='+', help='対象とするアプリID（指定したIDのみ処理）')
    all_parser.add_argument('--not-id', type=int, nargs='+', help='除外するアプリID（指定したID以外を処理）')

    # 出力ファイル一覧表示コマンド
    subparsers.add_parser(
        'outputs', parents=[common_parser],
        help='生成されるExcel/CSV/TSVファイルの一覧と概要を表示')

    return parser, group_parser


def main():
    """メイン関数"""
    parser, group_parser = build_parser()

    # 引数がない場合はヘルプと出力ファイル情報を表示
    if len(sys.argv) == 1:
        parser.print_help()
        print("\n")
        display_output_info()
        sys.exit(0)

    args = parser.parse_args()
    verbose = getattr(args, 'verbose', False)
    env_arg = getattr(args, 'env', None)

    # 出力ファイル一覧表示の場合（ログは初期化しない）
    if args.command == 'outputs':
        display_output_info()
        sys.exit(0)

    # group はアクション未指定なら使い方を表示して終了
    if args.command == 'group' and not args.action:
        group_parser.print_help()
        sys.exit(2)

    # ロギングの設定
    logger, log_file = setup_logging(verbose=verbose)
    logger.info("KintoneRunnerを起動しました")

    failures = 0
    exit_code = 0

    # ディレクトリの準備（allコマンドの場合のみ実行）
    if args.command == 'all':
        logger.info("ディレクトリの準備を開始します")
        try:
            prepare_directories()
        except Exception as e:
            logger.error(f"ディレクトリの準備中にエラーが発生しました: {e}")
            if "~$" in str(e):
                logger.error("Excelファイルが開かれているため処理を終了します。")
                fail("Excelファイルが開かれているため処理を続行できません。")
                note("Excelファイルを閉じてから再実行してください。")
                sys.exit(1)
            else:
                logger.warning("エラーが発生しましたが、処理を続行します。一部のファイルが正しく処理されない可能性があります。")
                warn(f"ディレクトリの準備中にエラーが発生しました: {e}")
                note("処理を続行しますが、一部のファイルが正しく処理されない可能性があります。")
    # appコマンドの場合、特定のアプリIDのディレクトリのみ準備
    elif args.command == 'app' and args.id:
        logger.info(f"アプリID {args.id} のディレクトリ準備を開始します")
        try:
            prepare_app_directories(args.id)
        except Exception as e:
            logger.error(f"ディレクトリの準備中にエラーが発生しました: {e}")
            if "~$" in str(e):
                logger.error("Excelファイルが開かれているため処理を終了します。")
                fail("Excelファイルが開かれているため処理を続行できません。")
                note("Excelファイルを閉じてから再実行してください。")
                sys.exit(1)
            else:
                logger.warning("エラーが発生しましたが、処理を続行します。一部のファイルが正しく処理されない可能性があります。")
                warn(f"ディレクトリの準備中にエラーが発生しました: {e}")
                note("処理を続行しますが、一部のファイルが正しく処理されない可能性があります。")

    # 最低限のディレクトリ作成を確保
    OUTPUT_DIR.mkdir(exist_ok=True)
    PREVIOUS_OUTPUT_DIR.mkdir(exist_ok=True)
    BACKUP_DIR.mkdir(exist_ok=True)

    # 設定ファイルの読み込み
    env_file = Path(env_arg) if env_arg else ENV_FILE
    config = load_env_config(env_file)
    if 'app_tokens' in config:
        # アプリIDのフィルタリング
        if args.command == 'all':
            if getattr(args, 'id', None):
                # --id が指定された場合、指定されたIDのみを対象とする
                target_ids = [str(i) for i in args.id]
                config['app_tokens'] = {k: v for k, v in config['app_tokens'].items() if str(k) in target_ids}
            elif getattr(args, 'not_id', None):
                # --not-id が指定された場合、指定されたID以外を対象とする
                exclude_ids = [str(i) for i in args.not_id]
                config['app_tokens'] = {k: v for k, v in config['app_tokens'].items() if str(k) not in exclude_ids}
    logger.info(f"設定ファイル {env_file} を読み込みました")

    # コマンドに応じて処理を実行
    if args.command == 'users':
        heading("ユーザーとグループ情報の取得")
        _, failed = get_user_group_info(config, logger, args.format)
        failures += failed

    elif args.command == 'app':
        heading("アプリ設定のダウンロード")
        status, failed = get_app_json(config, logger, args.id)
        failures += failed
        if status == "forbidden":
            fail("APIトークンの権限が不足しています。レコード閲覧を付けてアプリを更新し、再試行してください。")
            note(f"ログ: {rel_path(log_file)}")
            sys.exit(3)
        if not failed:
            # appコマンドで特定のアプリIDが指定された場合、事後処理も実行
            if args.id:
                backup_dir = backup_output()
                logger.info(f"出力ファイルを {backup_dir} にバックアップしました")
                note(f"バックアップ: {rel_path(backup_dir)}")

                # ファイル名から日時部分を除去
                remove_datetime_suffix(OUTPUT_DIR)

    elif args.command == 'acl':
        heading("ACL情報のExcel変換")
        warn_if_no_group_master()
        _, failed = generate_acl_excel(config, logger, args.id)
        failures += failed

    elif args.command == 'summary':
        heading("アプリ設定一覧表の生成")
        _, failed = generate_summary_excel(config, logger, args.output)
        failures += failed
        if failed:
            print_result_summary(log_file, failures)
            logger.info("KintoneRunnerを終了します")
            sys.exit(1)

    elif args.command == 'group':
        heading(f"グループ操作: {args.action}")
        if args.action == 'list':
            result, failed = manage_groups(config, logger, 'list', env_file=env_file)
            failures += failed
            if not failed:
                if result and result.strip():
                    print(result.rstrip())
                ok("グループ一覧を表示しました")

        elif args.action == 'search':
            result, failed = manage_groups(config, logger, 'search', {'keyword': args.keyword}, env_file=env_file)
            failures += failed
            if not failed:
                if result and result.strip():
                    print(result.rstrip())
                ok(f"'{args.keyword}' の検索結果を表示しました")

        elif args.action == 'add':
            result, failed = manage_groups(config, logger, 'add', {'user': args.user, 'group': args.group}, env_file=env_file)
            failures += failed
            if not failed:
                if result and result.strip():
                    print(result.rstrip())
                ok(f"ユーザー {args.user} をグループ {args.group} に追加しました")

        elif args.action == 'remove':
            result, failed = manage_groups(config, logger, 'remove', {'user': args.user}, env_file=env_file)
            failures += failed
            if not failed:
                if result and result.strip():
                    print(result.rstrip())
                ok(f"ユーザー {args.user} をグループから削除しました")

    elif args.command == 'notifications':
        heading("通知設定のExcel変換")
        warn_if_no_group_master()
        _, failed = generate_notifications_excel(config, logger, args.id)
        failures += failed

    elif args.command == 'process_workflow':
        heading("プロセスワークフローのExcel変換")
        _, failed = generate_process_workflow_excel(config, logger, args.id)
        failures += failed

    elif args.command == 'search':
        failures += command_search(
            config, logger,
            keywords=args.keywords,
            app_id=args.id,
            no_excel=args.no_excel,
            verbose=verbose,
        )

    elif args.command == 'webhooks':
        failures += command_webhooks(config, logger, args.id, verbose=verbose)

    elif args.command == 'urls':
        failures += command_urls(
            config, logger,
            app_ids=args.id,
            output=args.output,
            verbose=verbose,
        )

    elif args.command == 'all':
        # すべての機能を順番に実行
        logger.info("すべての機能を順番に実行します")

        # 1. ユーザーとグループ情報の取得
        heading("[1/6] ユーザーとグループ情報の取得")
        _, failed = get_user_group_info(config, logger)
        failures += failed

        # 2. アプリのJSONデータ取得
        heading("[2/6] アプリ設定のダウンロード")
        status, failed = get_app_json(config, logger)
        failures += failed
        if status == "forbidden":
            fail("APIトークンの権限が不足しています。レコード閲覧を付けてアプリを更新し、再試行してください。")
            note(f"ログ: {rel_path(log_file)}")
            sys.exit(3)

        # 3. ACL情報のExcel変換
        heading("[3/6] ACL情報のExcel変換")
        warn_if_no_group_master()
        _, failed = generate_acl_excel(config, logger)
        failures += failed

        # 4. アプリ設定一覧表の生成
        heading("[4/6] アプリ設定一覧表の生成")
        _, failed = generate_summary_excel(config, logger)
        failures += failed

        # 5. 通知設定のExcel変換
        heading("[5/6] 通知設定のExcel変換")
        warn_if_no_group_master()
        _, failed = generate_notifications_excel(config, logger)
        failures += failed

        # 6. プロセスワークフローのExcel変換
        heading("[6/6] プロセスワークフローのExcel変換")
        _, failed = generate_process_workflow_excel(config, logger)
        failures += failed

        # 処理完了後にバックアップを作成
        backup_dir = backup_output()
        logger.info(f"出力ファイルを {backup_dir} にバックアップしました")
        note(f"バックアップ: {rel_path(backup_dir)}")

        # ファイル名から日時部分を除去
        remove_datetime_suffix(OUTPUT_DIR)

    print_result_summary(log_file, failures)
    logger.info("KintoneRunnerを終了します")

    if failures:
        exit_code = 1
    sys.exit(exit_code)


if __name__ == "__main__":
    main()
