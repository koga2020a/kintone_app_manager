#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
kintoneアプリの通知設定をExcelに出力するスクリプト
"""

import os
import sys
import yaml
import json
import argparse
import logging
from pathlib import Path
from datetime import datetime
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# 定数定義
SCRIPT_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = SCRIPT_DIR.parent / "output"

def setup_logging():
    """ロギングの設定"""
    log_dir = SCRIPT_DIR.parent / "logs"
    log_dir.mkdir(exist_ok=True)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = log_dir / f"notifications_to_excel_{timestamp}.log"
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()
        ]
    )
    
    return logging.getLogger("notifications_to_excel")

def load_yaml_file(file_path):
    """YAMLファイルを読み込む"""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return yaml.safe_load(f)
    except Exception as e:
        raise Exception(f"YAMLファイルの読み込みに失敗しました: {file_path} - {e}")

def find_app_directory(base_dir, app_id):
    """アプリIDに対応するディレクトリを探す"""
    return next((d for d in base_dir.iterdir() if d.is_dir() and d.name.startswith(f"{app_id}_")), None)

def find_latest_group_user_excel(output_dir):
    """最新のユーザー・グループExcelファイルを探す"""
    excel_files = list(output_dir.glob("kintone_users_groups_*.xlsx"))
    
    if not excel_files:
        return None
    
    # 日時でソートして最新のファイルを返す
    latest_file = max(excel_files, key=lambda f: f.stat().st_mtime)
    return latest_file

def load_group_yaml_data(yaml_path):
    """
    group_user_list.yaml からグループとユーザー情報を読み込む
    
    Args:
        yaml_path: group_user_list.yaml のファイルパス
        
    Returns:
        dict: グループ情報の辞書
    """
    try:
        if not yaml_path or not yaml_path.exists():
            logging.warning(f"グループユーザーリストファイルが見つかりません: {yaml_path}")
            return {}
            
        with open(yaml_path, 'r', encoding='utf-8') as f:
            group_data = yaml.safe_load(f)
            logging.info(f"group_user_list.yaml から {len(group_data)} 件のグループ情報を読み込みました。")
            return group_data
    except Exception as e:
        logging.warning(f"グループユーザーリストの読み込みに失敗しました: {e}")
        return {}

# もう一つのデータソースからグループ情報を読み込む関数を追加
def load_group_list_yaml(yaml_dir):
    """
    グループコードとグループ名のマッピングを読み込む
    """
    try:
        group_list_path = yaml_dir / "group_list.yaml"
        if not group_list_path.exists():
            logging.warning(f"グループリストファイルが見つかりません: {group_list_path}")
            return {}
            
        with open(group_list_path, 'r', encoding='utf-8') as f:
            group_mapping = yaml.safe_load(f)
            return group_mapping
    except Exception as e:
        logging.warning(f"グループリストの読み込みに失敗しました: {e}")
        return {}

# ユーザー情報を読み込む関数を追加
def load_user_list_yaml(yaml_dir):
    """
    ユーザー情報を読み込む
    """
    try:
        user_list_path = yaml_dir / "user_list.yaml"
        if not user_list_path.exists():
            logging.warning(f"ユーザーリストファイルが見つかりません: {user_list_path}")
            return {}
            
        with open(user_list_path, 'r', encoding='utf-8') as f:
            user_data = yaml.safe_load(f)
            return user_data
    except Exception as e:
        logging.warning(f"ユーザーリストの読み込みに失敗しました: {e}")
        return {}

def find_group_user_list_yaml():
    """group_user_list.yamlファイルを探す"""
    # 検索場所のリスト
    search_paths = [
        Path(__file__).resolve().parent.parent,  # プロジェクトルート
        Path(__file__).resolve().parent,  # notifications_to_excel.pyと同じディレクトリ
        OUTPUT_DIR,  # 出力ディレクトリ
        Path.cwd()  # カレントディレクトリ
    ]
    
    for path in search_paths:
        yaml_path = path / "group_user_list.yaml"
        if yaml_path.exists():
            logging.info(f"group_user_list.yaml が見つかりました: {yaml_path}")
            return yaml_path
    
    logging.warning("group_user_list.yaml が見つかりませんでした")
    return None

def create_notification_excel(app_id, general_data, record_data, reminder_data, output_file=None):
    """通知設定をExcelに出力する"""
    
    if output_file is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = OUTPUT_DIR / f"{app_id}_notifications_{timestamp}.xlsx"
    
    # Excelワークブックを作成
    wb = Workbook()
    
    # スタイル定義
    header_font = Font(bold=True, size=11)
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    header_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    # 罫線スタイル
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # group_user_list.yamlからグループ情報を読み込む（マスターデータとして利用）
    group_yaml_path = find_group_user_list_yaml()
    group_yaml_data = load_group_yaml_data(group_yaml_path)
    
    # 収集したグループコードのリスト（シート下部にグループ情報表示用）
    collected_group_codes = []
    
    # 1. 一般通知設定のシート作成
    if general_data:
        create_general_notifications_sheet(wb, general_data, header_font, header_fill, header_alignment, thin_border, group_yaml_data, collected_group_codes)
        
        # 一般通知シートにグループ情報テーブル追加
        general_sheet = wb.active
        add_group_info_table(general_sheet, group_yaml_data, collected_group_codes, header_font, header_fill, header_alignment, thin_border)
    
    # 2. レコード通知設定のシート作成
    if record_data:
        create_record_notifications_sheet(wb, record_data, header_font, header_fill, header_alignment, thin_border, group_yaml_data, collected_group_codes)
        
        # レコード通知シートにグループ情報テーブル追加
        record_sheet = wb["レコード通知設定"]
        add_group_info_table(record_sheet, group_yaml_data, collected_group_codes, header_font, header_fill, header_alignment, thin_border)
    
    # 3. リマインダー通知設定のシート作成
    if reminder_data:
        create_reminder_notifications_sheet(wb, reminder_data, header_font, header_fill, header_alignment, thin_border, group_yaml_data, collected_group_codes)
        
        # リマインダー通知シートにグループ情報テーブル追加
        reminder_sheet = wb["リマインダー通知設定"]
        add_group_info_table(reminder_sheet, group_yaml_data, collected_group_codes, header_font, header_fill, header_alignment, thin_border)
    
    # Excelファイルを保存
    wb.save(output_file)
    logging.info(f"通知設定をExcelに出力しました: {output_file}")
    
    return output_file

def add_group_members_table(ws, row_idx, group_codes, header_font, header_fill, header_alignment, thin_border, group_yaml_data, collected_group_codes):
    """グループメンバー情報の表を追加"""

    if not group_codes or not group_yaml_data:
        return row_idx
    
    # グループ情報の見出し
    row_idx += 2
    ws.cell(row=row_idx, column=1).value = "グループメンバー情報"
    ws.cell(row=row_idx, column=1).font = Font(bold=True, size=12)
    row_idx += 1
    
    # 重複するグループコードを除去
    unique_group_codes = list(set(group_codes))
    
    for group_code in unique_group_codes:
        # グループが存在しない場合はスキップ
        if group_code not in group_yaml_data:
            logging.warning(f"グループ {group_code} の情報が見つかりません")
            continue
        
        group_info = group_yaml_data[group_code]
        group_name = group_info.get('name', '不明なグループ')
        members = group_info.get('users', [])
        
        # グループ名の行
        ws.cell(row=row_idx, column=1).value = f"グループ: {group_name} ({group_code})"
        ws.cell(row=row_idx, column=1).font = Font(bold=True)
        ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=3)
        row_idx += 1
        
        # ヘッダー行
        headers = ["No.", "ユーザー名", "メールアドレス"]
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.value = header
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
            cell.border = thin_border
        row_idx += 1
        
        # メンバー行
        for i, user in enumerate(members, 1):
            row_data = [
                i,
                user.get('username', '不明'),
                user.get('email', '')
            ]
            
            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.value = value
                cell.border = thin_border
            
            row_idx += 1
        
        # グループ間の空白
        row_idx += 1
    
    # 列幅の調整
    ws.column_dimensions["A"].width = 5    # No.
    ws.column_dimensions["B"].width = 25   # ユーザー名
    ws.column_dimensions["C"].width = 30   # メールアドレス
    
    return row_idx

def create_general_notifications_sheet(wb, data, header_font, header_fill, header_alignment, thin_border, group_yaml_data, collected_group_codes):
    """一般通知設定のシートを作成"""
    ws = wb.create_sheet(title="一般通知設定")
    
    # ヘッダー行 - フィールドタイプ列を追加
    headers = ["No.", "通知先タイプ", "フィールドタイプ", "通知先", "サブグループ含む", "レコード追加", "レコード編集", "コメント追加", "ステータス変更", "ファイル読込"]
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.value = header
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border
    
    # データ行
    notifications = data.get("notifications", [])
    
    # データ行の背景色を交互に設定
    light_blue_fill = PatternFill(start_color="EBF1F5", end_color="EBF1F5", fill_type="solid")
    
    # グループの通知先を収集
    group_codes = []
    
    for row_idx, notify in enumerate(notifications, 2):
        entity = notify.get("entity", {})
        entity_type = entity.get("type", "")
        entity_code = entity.get("code", "")
        
        # グループコードを収集
        if entity_type == "GROUP":
            group_codes.append(entity_code)
        
        # 通知先タイプを日本語に変換
        type_jp = ""
        field_type = ""
        if entity_type == "USER":
            type_jp = "ユーザー"
        elif entity_type == "GROUP":
            type_jp = "グループ"
        elif entity_type == "ORGANIZATION":
            type_jp = "組織"
        elif entity_type == "FIELD_ENTITY":
            type_jp = "フィールド"
            # フィールドタイプの取得
            if "type" in entity:
                if entity.get("type") == "CREATOR":
                    field_type = "作成者"
                elif entity.get("type") == "MODIFIER":
                    field_type = "更新者"
                elif entity.get("type") == "USER_SELECT":
                    field_type = "ユーザー選択"
                elif entity.get("type") == "GROUP_SELECT":
                    field_type = "グループ選択"
                elif entity.get("type") == "ORGANIZATION_SELECT":
                    field_type = "組織選択"
                else:
                    field_type = entity.get("type", "")
        else:
            type_jp = entity_type
        
        # データを行に設定
        row_data = [
            row_idx - 1,  # No.
            type_jp,  # 通知先タイプ
            field_type,  # フィールドタイプ - 新しい列
            entity_code,  # 通知先
            "はい" if notify.get("includeSubs", False) else "いいえ",  # サブグループ含む
            "はい" if notify.get("recordAdded", False) else "いいえ",  # レコード追加
            "はい" if notify.get("recordEdited", False) else "いいえ",  # レコード編集
            "はい" if notify.get("commentAdded", False) else "いいえ",  # コメント追加
            "はい" if notify.get("statusChanged", False) else "いいえ",  # ステータス変更
            "はい" if notify.get("fileImported", False) else "いいえ",  # ファイル読込
        ]
        
        # 行の背景色を交互に設定
        row_fill = light_blue_fill if row_idx % 2 == 0 else None
        
        for col_idx, value in enumerate(row_data, 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.value = value
            cell.border = thin_border
            if row_fill:
                cell.fill = row_fill
            if col_idx >= 5:  # チェックボックス的な列は中央揃え
                cell.alignment = Alignment(horizontal='center')
    
    # コメント通知設定
    row_idx = len(notifications) + 3
    ws.cell(row=row_idx, column=1).value = "コメント投稿者への通知:"
    ws.cell(row=row_idx, column=1).font = Font(bold=True)
    ws.cell(row=row_idx, column=2).value = "はい" if data.get("notifyToCommenter", False) else "いいえ"
    ws.cell(row=row_idx, column=2).alignment = Alignment(horizontal='center')
    
    # セクション背景色
    section_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    for col_idx in range(1, 3):
        cell = ws.cell(row=row_idx, column=col_idx)
        cell.fill = section_fill
        cell.border = thin_border
    
    # 列幅の調整
    for col_idx in range(1, len(headers) + 1):
        column_letter = get_column_letter(col_idx)
        if col_idx == 1:  # No.列
            ws.column_dimensions[column_letter].width = 5
        elif col_idx == 2:  # 通知先タイプ列
            ws.column_dimensions[column_letter].width = 12
        elif col_idx == 3:  # フィールドタイプ列
            ws.column_dimensions[column_letter].width = 15
        elif col_idx == 4:  # 通知先列
            ws.column_dimensions[column_letter].width = 20
        else:  # その他の列
            ws.column_dimensions[column_letter].width = 15
    
    # グループメンバー情報を追加
    if group_codes:
        row_idx = add_group_members_table(ws, row_idx, group_codes, header_font, header_fill, header_alignment, thin_border, group_yaml_data, collected_group_codes)

def create_record_notifications_sheet(wb, data, header_font, header_fill, header_alignment, thin_border, group_yaml_data, collected_group_codes):
    """レコード通知設定のシートを作成"""
    ws = wb.create_sheet(title="レコード通知設定")
    
    # ヘッダー行 - フィールドタイプ列を追加
    headers = ["No.", "通知先タイプ", "フィールドタイプ", "通知先", "通知条件", "条件内容"]
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.value = header
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border
    
    # データ行
    row_idx = 2
    notifications = data.get("notifications", [])
    
    # データ行の背景色を設定
    light_blue_fill = PatternFill(start_color="EBF1F5", end_color="EBF1F5", fill_type="solid")
    light_green_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    
    # グループの通知先を収集
    group_codes = []
    
    for notify_idx, notify in enumerate(notifications, 1):
        entity = notify.get("entity", {})
        entity_type = entity.get("type", "")
        entity_code = entity.get("code", "")
        
        # グループコードを収集
        if entity_type == "GROUP":
            group_codes.append(entity_code)
        
        # 通知先タイプを日本語に変換
        type_jp = ""
        field_type = ""
        if entity_type == "USER":
            type_jp = "ユーザー"
        elif entity_type == "GROUP":
            type_jp = "グループ"
        elif entity_type == "ORGANIZATION":
            type_jp = "組織"
        elif entity_type == "FIELD_ENTITY":
            type_jp = "フィールド"
            # フィールドタイプの取得
            if "type" in entity:
                if entity.get("type") == "CREATOR":
                    field_type = "作成者"
                elif entity.get("type") == "MODIFIER":
                    field_type = "更新者"
                elif entity.get("type") == "USER_SELECT":
                    field_type = "ユーザー選択"
                elif entity.get("type") == "GROUP_SELECT":
                    field_type = "グループ選択"
                elif entity.get("type") == "ORGANIZATION_SELECT":
                    field_type = "組織選択"
                else:
                    field_type = entity.get("type", "")
        else:
            type_jp = entity_type
        
        # 通知条件
        conditions = notify.get("condition", {}).get("conditions", [])
        
        # 通知先ごとに背景色を交互に変更
        notify_fill = light_blue_fill if notify_idx % 2 == 1 else light_green_fill
        
        if not conditions:
            # 条件がない場合は1行だけ出力
            row_data = [
                notify_idx,  # No.
                type_jp,  # 通知先タイプ
                field_type,  # フィールドタイプ - 新しい列
                entity_code,  # 通知先
                "条件なし",  # 通知条件
                "",  # 条件内容
            ]
            
            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.value = value
                cell.border = thin_border
                cell.fill = notify_fill
            
            row_idx += 1
        else:
            # 条件ごとに行を作成
            for cond_idx, condition in enumerate(conditions):
                cond_type = condition.get("type", "")
                field_code = condition.get("field", {}).get("code", "")
                operator = condition.get("operator", "")
                value = condition.get("value", "")
                
                # 条件タイプを日本語に変換
                if cond_type == "CONDITION":
                    cond_type_jp = "フィールド条件"
                elif cond_type == "STATUS":
                    cond_type_jp = "ステータス条件"
                else:
                    cond_type_jp = cond_type
                
                # 演算子を日本語に変換
                operator_jp = operator
                if operator == "=":
                    operator_jp = "等しい"
                elif operator == "!=":
                    operator_jp = "等しくない"
                elif operator == ">":
                    operator_jp = "より大きい"
                elif operator == "<":
                    operator_jp = "より小さい"
                elif operator == ">=":
                    operator_jp = "以上"
                elif operator == "<=":
                    operator_jp = "以下"
                elif operator == "in":
                    operator_jp = "含む"
                elif operator == "not in":
                    operator_jp = "含まない"
                
                # 条件内容を整形
                condition_content = f"{field_code} {operator_jp} {value}"
                
                row_data = [
                    notify_idx if cond_idx == 0 else "",  # No.（最初の条件の行のみ表示）
                    type_jp if cond_idx == 0 else "",  # 通知先タイプ（最初の条件の行のみ表示）
                    field_type if cond_idx == 0 else "",  # フィールドタイプ（最初の条件の行のみ表示）
                    entity_code if cond_idx == 0 else "",  # 通知先（最初の条件の行のみ表示）
                    cond_type_jp,  # 通知条件
                    condition_content,  # 条件内容
                ]
                
                for col_idx, value in enumerate(row_data, 1):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    cell.value = value
                    cell.border = thin_border
                    cell.fill = notify_fill
                
                row_idx += 1
    
    # 列幅の調整
    for col_idx in range(1, len(headers) + 1):
        column_letter = get_column_letter(col_idx)
        if col_idx == 1:  # No.列
            ws.column_dimensions[column_letter].width = 5
        elif col_idx == 2:  # 通知先タイプ列
            ws.column_dimensions[column_letter].width = 12
        elif col_idx == 3:  # フィールドタイプ列
            ws.column_dimensions[column_letter].width = 15
        elif col_idx == 4:  # 通知先列
            ws.column_dimensions[column_letter].width = 20
        elif col_idx == 5:  # 通知条件列
            ws.column_dimensions[column_letter].width = 15
        else:  # 条件内容列
            ws.column_dimensions[column_letter].width = 40
    
    # グループメンバー情報を追加
    if group_codes:
        row_idx = add_group_members_table(ws, row_idx, group_codes, header_font, header_fill, header_alignment, thin_border, group_yaml_data, collected_group_codes)

def create_reminder_notifications_sheet(wb, data, header_font, header_fill, header_alignment, thin_border, group_yaml_data, collected_group_codes):
    """リマインダー通知設定のシートを作成"""
    ws = wb.create_sheet(title="リマインダー通知設定")
    
    # ヘッダー行 - フィールドタイプ列を追加
    headers = ["No.", "リマインダー名", "通知先タイプ", "フィールドタイプ", "通知先", "日時フィールド", "条件", "通知タイミング"]
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.value = header
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border
    
    # データ行
    row_idx = 2
    reminders = data.get("reminders", [])
    
    # データ行の背景色を設定
    light_blue_fill = PatternFill(start_color="EBF1F5", end_color="EBF1F5", fill_type="solid")
    light_yellow_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    
    # グループの通知先を収集
    group_codes = []
    
    for remind_idx, remind in enumerate(reminders, 1):
        name = remind.get("title", "")
        date_field = remind.get("timing", {}).get("field", {}).get("code", "")
        
        # 通知タイミング
        timing_type = remind.get("timing", {}).get("type", "")
        timing_value = remind.get("timing", {}).get("value", "")
        timing_unit = remind.get("timing", {}).get("unit", "")
        
        # 通知タイミングを整形
        timing_jp = ""
        if timing_type == "BEFORE":
            timing_jp = f"{timing_value}{timing_unit}前"
        elif timing_type == "AFTER":
            timing_jp = f"{timing_value}{timing_unit}後"
        else:
            timing_jp = f"{timing_type}: {timing_value} {timing_unit}"
        
        # 条件を整形
        condition_type = remind.get("filterCond", "")
        condition_jp = "全レコード" if not condition_type else f"条件式: {condition_type}"
        
        # リマインダーごとに背景色を交互に変更
        remind_fill = light_blue_fill if remind_idx % 2 == 1 else light_yellow_fill
        
        # 通知先
        recipients = remind.get("recipients", [])
        
        if not recipients:
            # 通知先がない場合は1行だけ出力
            row_data = [
                remind_idx,  # No.
                name,  # リマインダー名
                "",  # 通知先タイプ
                "",  # フィールドタイプ - 新しい列
                "通知先なし",  # 通知先
                date_field,  # 日時フィールド
                condition_jp,  # 条件
                timing_jp,  # 通知タイミング
            ]
            
            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.value = value
                cell.border = thin_border
                cell.fill = remind_fill
            
            row_idx += 1
        else:
            # 通知先ごとに行を作成
            for recip_idx, recipient in enumerate(recipients):
                entity = recipient.get("entity", {})
                entity_type = entity.get("type", "")
                entity_code = entity.get("code", "")
                
                # グループコードを収集
                if entity_type == "GROUP":
                    group_codes.append(entity_code)
                
                # 通知先タイプを日本語に変換
                type_jp = ""
                field_type = ""
                if entity_type == "USER":
                    type_jp = "ユーザー"
                elif entity_type == "GROUP":
                    type_jp = "グループ"
                elif entity_type == "ORGANIZATION":
                    type_jp = "組織"
                elif entity_type == "FIELD_ENTITY":
                    type_jp = "フィールド"
                    # フィールドタイプの取得
                    if "type" in entity:
                        if entity.get("type") == "CREATOR":
                            field_type = "作成者"
                        elif entity.get("type") == "MODIFIER":
                            field_type = "更新者"
                        elif entity.get("type") == "USER_SELECT":
                            field_type = "ユーザー選択"
                        elif entity.get("type") == "GROUP_SELECT":
                            field_type = "グループ選択"
                        elif entity.get("type") == "ORGANIZATION_SELECT":
                            field_type = "組織選択"
                        else:
                            field_type = entity.get("type", "")
                else:
                    type_jp = entity_type
                
                row_data = [
                    remind_idx if recip_idx == 0 else "",  # No.（最初の通知先の行のみ表示）
                    name if recip_idx == 0 else "",  # リマインダー名（最初の通知先の行のみ表示）
                    type_jp,  # 通知先タイプ
                    field_type,  # フィールドタイプ - 新しい列
                    entity_code,  # 通知先
                    date_field if recip_idx == 0 else "",  # 日時フィールド（最初の通知先の行のみ表示）
                    condition_jp if recip_idx == 0 else "",  # 条件（最初の通知先の行のみ表示）
                    timing_jp if recip_idx == 0 else "",  # 通知タイミング（最初の通知先の行のみ表示）
                ]
                
                for col_idx, value in enumerate(row_data, 1):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    cell.value = value
                    cell.border = thin_border
                    cell.fill = remind_fill
                
                row_idx += 1
    
    # 列幅の調整
    for col_idx in range(1, len(headers) + 1):
        column_letter = get_column_letter(col_idx)
        if col_idx == 1:  # No.列
            ws.column_dimensions[column_letter].width = 5
        elif col_idx == 2:  # リマインダー名列
            ws.column_dimensions[column_letter].width = 20
        elif col_idx == 3:  # 通知先タイプ列
            ws.column_dimensions[column_letter].width = 12
        elif col_idx == 4:  # フィールドタイプ列
            ws.column_dimensions[column_letter].width = 15
        elif col_idx == 5:  # 通知先列
            ws.column_dimensions[column_letter].width = 20
        elif col_idx == 6:  # 日時フィールド列
            ws.column_dimensions[column_letter].width = 15
        elif col_idx == 7:  # 条件列
            ws.column_dimensions[column_letter].width = 30
        else:  # 通知タイミング列
            ws.column_dimensions[column_letter].width = 15

    # グループメンバー情報を追加
    if group_codes:
        row_idx = add_group_members_table(ws, row_idx, group_codes, header_font, header_fill, header_alignment, thin_border, group_yaml_data, collected_group_codes)

def add_group_info_table(ws, group_yaml_data, collected_group_codes, header_font, header_fill, header_alignment, thin_border):
    """シートの下部にグループ情報テーブルを追加する"""
    if not collected_group_codes or not group_yaml_data:
        return
    
    # 現在のデータの最終行を取得
    max_row = ws.max_row
    
    # 2行空けてからテーブルを開始
    start_row = max_row + 3
    
    # タイトル行
    ws.cell(row=start_row, column=1, value="グループ情報一覧").font = Font(bold=True, size=14)
    start_row += 2
    
    # 各グループについてテーブルを作成
    for group_code in collected_group_codes:
        if group_code not in group_yaml_data:
            logging.warning(f"グループコード {group_code} が group_user_list.yaml に見つかりません。")
            continue
            
        group_info = group_yaml_data[group_code]
        group_name = group_info.get('name', group_code)
        
        # グループヘッダー
        ws.cell(row=start_row, column=1, value=f"グループ名: {group_name} (コード: {group_code})").font = Font(bold=True, size=12)
        start_row += 1
        
        # ユーザーテーブルのヘッダー
        headers = ["ユーザー名", "メールアドレス"]
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=start_row, column=col, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
            cell.border = thin_border
        
        start_row += 1
        
        # ユーザーデータの行
        users = group_info.get('users', [])
        if not users:
            ws.cell(row=start_row, column=1, value="(ユーザーなし)").border = thin_border
            ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=2)
            start_row += 1
        else:
            for user in users:
                username = user.get('username', '')
                email = user.get('email', '')
                
                ws.cell(row=start_row, column=1, value=username).border = thin_border
                ws.cell(row=start_row, column=2, value=email).border = thin_border
                start_row += 1
        
        # グループ間に空行を入れる
        start_row += 2

def main():
    """メイン関数"""
    # コマンドライン引数の解析
    parser = argparse.ArgumentParser(description='kintoneアプリの通知設定をExcelに出力するスクリプト')
    parser.add_argument('app_id', type=int, help='アプリID')
    parser.add_argument('--output', type=str, help='出力ファイル名')
    
    args = parser.parse_args()
    
    # ロギングの設定
    logger = setup_logging()
    logger.info(f"アプリID {args.app_id} の通知設定のExcel出力を開始します")
    
    try:
        # アプリIDに対応するディレクトリを探す
        app_dir = find_app_directory(OUTPUT_DIR, args.app_id)
        if not app_dir:
            logger.error(f"アプリID {args.app_id} に対応するディレクトリが見つかりません")
            print(f"エラー: アプリID {args.app_id} に対応するディレクトリが見つかりません")
            sys.exit(1)
        
        # 通知設定ファイルのパス
        general_file = app_dir / f"{args.app_id}_general_notifications.yaml"
        record_file = app_dir / f"{args.app_id}_record_notifications.yaml"
        reminder_file = app_dir / f"{args.app_id}_reminder_notifications.yaml"
        
        # YAMLファイルの読み込み
        general_data = load_yaml_file(general_file) if general_file.exists() else None
        record_data = load_yaml_file(record_file) if record_file.exists() else None
        reminder_data = load_yaml_file(reminder_file) if reminder_file.exists() else None
        
        if not any([general_data, record_data, reminder_data]):
            logger.error(f"アプリID {args.app_id} の通知設定ファイルが見つかりません")
            print(f"エラー: アプリID {args.app_id} の通知設定ファイルが見つかりません")
            sys.exit(1)
        
        # 出力ファイル名
        output_file = args.output
        if not output_file:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = app_dir / f"{args.app_id}_notifications_{timestamp}.xlsx"
        else:
            output_file = Path(output_file)
        
        # Excelファイルの作成
        excel_file = create_notification_excel(args.app_id, general_data, record_data, reminder_data, output_file)
        
        logger.info(f"通知設定を {excel_file} に出力しました")
        print(f"通知設定を {excel_file} に出力しました")
        
    except Exception as e:
        logger.error(f"エラーが発生しました: {e}")
        import traceback
        logger.error(traceback.format_exc())
        print(f"エラー: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main() 