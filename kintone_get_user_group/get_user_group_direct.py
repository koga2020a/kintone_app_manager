#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Kintoneのユーザーとグループ情報を直接APIで取得するスクリプト
"""

import sys
import os
import argparse
import logging
from pathlib import Path
from datetime import datetime

# スクリプトの親ディレクトリを取得
SCRIPT_DIR = Path(__file__).resolve().parent
PARENT_DIR = SCRIPT_DIR.parent

# 親ディレクトリをPythonパスに追加
sys.path.append(str(PARENT_DIR))

# kintone_userlibのインポート
from lib.kintone_userlib.client import KintoneClient
from lib.kintone_userlib.manager import UserManager

def setup_logging():
    """ロギングの設定"""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler()
        ]
    )
    return logging.getLogger("get_user_group_direct")

def main():
    """メイン関数"""
    # コマンドライン引数の解析
    parser = argparse.ArgumentParser(description='Kintoneのユーザーとグループ情報を直接APIで取得')
    parser.add_argument('--subdomain', required=True, help='Kintoneのサブドメイン')
    parser.add_argument('--username', required=True, help='Kintoneのユーザー名')
    parser.add_argument('--password', required=True, help='Kintoneのパスワード')
    parser.add_argument('--output', required=True, help='出力ファイルのパス')
    
    args = parser.parse_args()
    
    # ロギングの設定
    logger = setup_logging()
    logger.info("ユーザーとグループ情報の直接取得を開始します")
    
    try:
        # Kintoneクライアントの初期化
        logger.info("認証情報を設定中...")
        client = KintoneClient(args.subdomain, args.username, args.password, logger)

        # データの取得
        logger.info("全ユーザーを取得中...")
        all_users = client.get_all_users()

        logger.info("全グループを取得中...")
        all_groups = client.get_all_groups()

        # UserManagerの初期化とデータの設定
        manager = UserManager()
        for user in all_users:
            manager.add_user(user)
        for group in all_groups:
            manager.add_group(group)

        # グループとユーザーの関連付け
        logger.info("グループとユーザーの関連付けを開始します...")
        for group in all_groups:
            users_in_group = client.get_users_in_group(group.code)
            for user in users_in_group:
                group.add_user(user)

        # pickleで保存
        logger.info(f"データを {args.output} に保存中...")
        manager.to_pickle(args.output)
        logger.info("データの保存が完了しました")

        return 0
    except Exception as e:
        logger.error(f"ユーザーとグループ情報の取得中にエラーが発生しました: {e}")
        return 1

if __name__ == "__main__":
    sys.exit(main()) 