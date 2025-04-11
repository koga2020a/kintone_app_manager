#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
kintone_users_groups.pickleファイルの内容を読み込んでデバッグ表示するスクリプト
"""

import sys
import os
from pathlib import Path
import re
from lib.kintone_userlib.manager import UserManager
from typing import Dict, List

def get_primary_domain_from_env():
    """環境設定ファイルから主要ドメインを取得"""
    env_file = Path(__file__).resolve().parent / ".kintone.env"
    primary_domain = ""
    
    if env_file.exists():
        with open(env_file, 'r', encoding='utf-8') as f:
            for line in f:
                # "user_domain: "値"" の形式を検索
                match = re.search(r'user_domain:\s*"([^"]+)"', line)
                if match:
                    primary_domain = match.group(1)
                    break
    
    return primary_domain

def main():
    # スクリプトのディレクトリを取得
    script_dir = Path(__file__).resolve().parent
    pickle_file = script_dir / "output" / "kintone_users_groups.pickle"

    if not pickle_file.exists():
        print(f"エラー: pickleファイルが見つかりません: {pickle_file}")
        sys.exit(1)

    try:
        # 主要ドメインを取得
        primary_domain = get_primary_domain_from_env()
        if primary_domain:
            print(f"主要ドメイン: {primary_domain}")
        else:
            print("警告: 主要ドメインが設定されていません")
            
        # pickleファイルからデータを読み込む
        manager = UserManager.from_pickle(str(pickle_file))
        
        # 主要ドメインを設定
        if primary_domain:
            manager.set_primary_domain(primary_domain)
        
        # UserManagerのインスタンスであることを確認
        if not isinstance(manager, UserManager):
            print("エラー: 読み込んだオブジェクトはUserManagerのインスタンスではありません")
            sys.exit(1)
            
        # 必要な属性が存在しない場合は初期化
        if not hasattr(manager, 'user_groups'):
            print("注意: user_groups属性が存在しないため初期化します")
            manager.user_groups = {}
            
        if not hasattr(manager, 'group_users'):
            print("注意: group_users属性が存在しないため初期化します")
            manager.group_users = {}
            # グループとユーザーの関連を再構築
            for group_code, group in manager.groups.items():
                manager.group_users[group_code] = group.users
        
        # データの整合性をチェック・修正
        print("ユーザーとグループの関連を再構築しています...")
        
        # 1. まずgroup_usersを使ってuser_groupsを再構築
        for group_code, users in manager.group_users.items():
            group = manager.groups.get(group_code)
            if not group:
                continue
                
            for user in users:
                user_code = user.username
                if user_code not in manager.user_groups:
                    manager.user_groups[user_code] = []
                    
                if group not in manager.user_groups[user_code]:
                    manager.user_groups[user_code].append(group)
        
        # 2. ユーザーオブジェクトのgroupsリストも更新
        for user_code, user in manager.users.items():
            if user_code in manager.user_groups:
                user.groups = manager.user_groups[user_code]
        
        # 3. グループオブジェクトのusersリストも更新
        for group_code, group in manager.groups.items():
            if group_code in manager.group_users:
                group.users = manager.group_users[group_code]

        # 通常のデバッグ表示
        manager.debug_display()
        
        # 詳細なグループ-ユーザー関連の表示
        manager.display_group_users_detail()
        
    except Exception as e:
        print(f"エラーが発生しました: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main() 