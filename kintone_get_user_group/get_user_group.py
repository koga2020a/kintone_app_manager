import argparse
import sys
import base64
from getpass import getpass
from typing import List, Dict, Any

import requests

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
from openpyxl.utils import column_index_from_string, get_column_letter
import logging
import yaml
import glob
import zipfile
import tempfile
import os

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

  def get_all_users(self) -> List[Dict[str, Any]]:
    return self._fetch_data('users', {}, 'users')

  def get_all_groups(self) -> List[Dict[str, Any]]:
    return self._fetch_data('groups', {}, 'groups')

  def get_users_in_group(self, group_code: str) -> List[Dict[str, Any]]:
    params = {'code': group_code}
    return self._fetch_data('group/users', params, 'users')

class DataProcessor:
  def __init__(self, users: List[Dict[str, Any]], groups: List[Dict[str, Any]], client: KintoneClient, logger: logging.Logger):
    self.users = users
    self.groups = groups
    self.client = client
    self.user_id_to_info: Dict[str, Dict[str, Any]] = {}
    self.logger = logger

  def filter_groups(self) -> List[Dict[str, Any]]:
    filtered = [group for group in self.groups if group.get('name') != 'Everyone' and group.get('code')]
    self.logger.info(f"「Everyone」を除外し、codeが存在するグループ数: {len(filtered)}")
    return filtered

  def organize_groups(self, filtered_groups: List[Dict[str, Any]]) -> List[str]:
    group_names = [group['name'] for group in filtered_groups]
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
      user_id = str(user.get('id'))
      self.user_id_to_info[user_id] = {
        'ユーザーID': user_id,
        'ステータス': '停止中' if not user.get('valid', True) else '',
        'ログイン名': user.get('code'),
        '氏名': user.get('name'),
        'メールアドレス': user.get('email'),
        '所属グループ一覧': [],
        '最終アクセス日': '',
        '経過日数': ''
      }
    self.logger.info(f"ユーザー情報をマッピングしました。総ユーザー数: {len(self.user_id_to_info)}")

  def populate_group_memberships(self, filtered_groups: List[Dict[str, Any]]):
    self.logger.info("各グループの所属ユーザーを取得中...")
    for group in filtered_groups:
      group_code = group.get('code')
      group_name = group.get('name')
      self.logger.info(f"グループ '{group_name}' ({group_code}) のユーザーを取得中...")
      users_in_group = self.client.get_users_in_group(group_code)
      self.logger.info(f"グループ '{group_name}' に所属するユーザー数: {len(users_in_group)}")
      for user in users_in_group:
        user_id = str(user.get('id'))
        if user_id in self.user_id_to_info:
          self.user_id_to_info[user_id]['所属グループ一覧'].append(group_name)
          self.logger.debug(f"ユーザーID {user_id} はグループ '{group_name}' に所属しています。")
        else:
          # ユーザーが全ユーザーリストに存在しない場合（稀）
          self.user_id_to_info[user_id] = {
            'ユーザーID': user_id,
            'ステータス': '',
            'ログイン名': user.get('code'),
            '氏名': user.get('name'),
            'メールアドレス': user.get('email'),
            '所属グループ一覧': [group_name],
            '最終アクセス日': '',
            '経過日数': ''
          }
          self.logger.debug(f"ユーザーID {user_id} はグループ '{group_name}' に所属しています（新規追加）。")
    self.logger.info("グループの所属ユーザー情報を更新しました。")

  def generate_dataframes(self, group_names: List[str]) -> Dict[str, pd.DataFrame]:
    self.logger.info("データフレームを作成中...")
    user_data_active = []
    user_data_stopped = []

    for user_info in self.user_id_to_info.values():
      login_name = user_info['ログイン名'] or ''  # Noneの場合は空文字に
      email = user_info['メールアドレス'] or ''   # Noneの場合は空文字に
      
      # 条件に基づいて「相違」列を設定
      if login_name and email:  # 両方とも値が存在する場合のみ比較
        if login_name != email:
          if login_name.lower() == email.lower():
            discrepancy = "大小相違"
          else:
            discrepancy = "相違"
        else:
          discrepancy = ""
      else:
        discrepancy = ""  # どちらかが空の場合は相違なしとする
      
      user_info['相違'] = discrepancy
      user_info['所属グループ一覧'] = ', '.join(user_info['所属グループ一覧'])
      if user_info['ステータス'] == '停止中':
        user_data_stopped.append(user_info)
      else:
        user_data_active.append(user_info)

    df_active = pd.DataFrame(user_data_active)
    df_stopped = pd.DataFrame(user_data_stopped)

    # グループごとの「●」をマークする列を追加
    for group in group_names:
      df_active[group] = df_active['所属グループ一覧'].apply(
        lambda x: '●' if group in [g.strip() for g in x.split(',')] else ''
      )
      df_stopped[group] = df_stopped['所属グループ一覧'].apply(
        lambda x: '●' if group in [g.strip() for g in x.split(',')] else ''
      )

    # 列の順序を設定（「相違」列をB列に挿入し、GとHを初期から含める）
    basic_columns = ['ユーザーID', '相違', 'ステータス', 'ログイン名', '氏名', 'メールアドレス', '最終アクセス日', '経過日数', '所属グループ一覧']
    group_columns = group_names
    columns_order = basic_columns + group_columns

    df_active = df_active[columns_order].sort_values(by=['所属グループ一覧'], ascending=False)
    df_stopped = df_stopped[columns_order].sort_values(by=['所属グループ一覧'], ascending=False)

    self.logger.info("データフレームの作成が完了しました。")
    return {'アクティブ': df_active, '停止中': df_stopped}

  def export_group_user_list(self, filtered_groups: List[Dict[str, Any]]):
    """グループとユーザーの関連をYAMLファイルとして出力"""
    self.logger.info("group_user_list.yaml と group_user_list_NoUse.yaml を生成中...")
    
    active_group_data = {}
    inactive_group_data = {}
    
    # まず全グループの基本情報を設定
    for group in filtered_groups:
      group_code = group.get('code')
      active_group_data[group_code] = {
        'name': group.get('name'),
        'users': []
      }
      inactive_group_data[group_code] = {
        'name': group.get('name'),
        'users': []
      }
      
      # グループ内のユーザー情報を取得
      users_in_group = self.client.get_users_in_group(group_code)
      for user in users_in_group:
        user_info = {
          'username': user.get('code'),
          'email': user.get('email'),
          'id': str(user.get('id'))
        }
        # ユーザーの状態を確認
        if user.get('valid', True):
          active_group_data[group_code]['users'].append(user_info)
        else:
          inactive_group_data[group_code]['users'].append(user_info)
    
    # Everyoneグループを追加
    active_everyone_users = []
    inactive_everyone_users = []
    for user in self.users:
      user_info = {
        'username': user.get('code'),
        'email': user.get('email'),
        'id': str(user.get('id'))
      }
      if user.get('valid', True):
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
      
      self.logger.info("group_user_list.yaml と group_user_list_NoUse.yaml の生成が完了しました。")
    except Exception as e:
      self.logger.error(f"YAMLファイルの生成中にエラーが発生しました: {e}")

class ExcelExporter:
  def __init__(self, dataframes: Dict[str, pd.DataFrame], group_names: List[str], output_file: str, logger: logging.Logger):
    self.dataframes = dataframes
    self.group_names = group_names
    self.output_file = output_file
    self.logger = logger
    self.group_data = {}  # グループ情報を保持する辞書を追加

  def prepare_group_data(self, client: KintoneClient):
    """グループごとのユーザー情報を準備"""
    self.logger.info("グループ情報シート用のデータを準備中...")
    
    for group in self.group_names:
        group_users = []
        for df in self.dataframes.values():
            # グループに所属するユーザーを抽出
            mask = df[group] == '●'
            users = df[mask][['ユーザーID', 'ログイン名', '氏名', 'メールアドレス', 'ステータス']].copy()
            
            if not users.empty:
                # メールアドレスを分解してソート用の列を作成
                users['domain'] = users['メールアドレス'].str.split('@').str[1]
                users['localpart'] = users['メールアドレス'].str.split('@').str[0]
                # ドメインでソート、次に@前のローカル部分でソート
                users = users.sort_values(['domain', 'localpart'])
                # 一時的なソート用列を削除
                users = users.drop(['domain', 'localpart'], axis=1)
                # 停止中のユーザーに「●」を追加
                users['停止中'] = users['ステータス'].apply(lambda x: '●' if x == '停止中' else '')
                group_users.append(users)
        
        if group_users:
            self.group_data[group] = pd.concat(group_users, ignore_index=True)
        else:
            self.group_data[group] = pd.DataFrame(columns=['ユーザーID', 'ログイン名', '氏名', 'メールアドレス', '停止中'])

  def export_to_excel(self):
    self.logger.info("Excelファイルに出力中...")
    
    with pd.ExcelWriter(self.output_file, engine='openpyxl') as writer:
      # 既存シート（アクティブ、停止中）の出力
      for sheet_name, df in self.dataframes.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)
      
      # グループ情報シートを新規作成
      if self.group_data:
        sheet_name = 'グループ情報'
        start_row = 0
        # シートを追加
        workbook = writer.book
        ws = workbook.create_sheet(title=sheet_name)
        
        for group_name, df in self.group_data.items():
          # --- 1. グループ名行 ---
          ws.cell(row=start_row+1, column=1, value="グループ: " + group_name)
          start_row += 1
          
          # --- 2. ヘッダー行 ---
          headers = ["ユーザーID", "ログイン名", "氏名", "メールアドレス", "停止中"]
          for col, header in enumerate(headers, 1):
            ws.cell(row=start_row+1, column=col, value=header)
          start_row += 1
          
          # --- 3. データ行 ---
          if not df.empty:
            for r_idx, row in df.iterrows():
              ws.cell(row=start_row+1, column=1, value=row['ユーザーID'])
              ws.cell(row=start_row+1, column=2, value=row['ログイン名'])
              ws.cell(row=start_row+1, column=3, value=row['氏名'])
              ws.cell(row=start_row+1, column=4, value=row['メールアドレス'])
              ws.cell(row=start_row+1, column=5, value=row['停止中'])
              start_row += 1
          else:
            # データがない場合は空行を出力
            ws.cell(row=start_row+1, column=1, value="(データなし)")
            start_row += 1
          
          # --- 4. セット間に空行を追加 ---
          start_row += 1
        
        writer.sheets[sheet_name] = ws
    
    self.logger.info(f"Excelファイル '{self.output_file}' を作成しました。")

  def format_excel(self):
    self.logger.info("Excelファイルのフォーマットを設定中...")
    
    wb = load_workbook(self.output_file)
    sheets = ['アクティブ', '停止中']

    # 単位変換：px → 文字数（openpyxlでは幅は文字数）
    def px_to_char(px):
        return px / 7

    # ヘッダー行の基本スタイル
    header_fill = PatternFill(start_color='243C5C', end_color='243C5C', fill_type='solid')
    header_font = Font(bold=True, color='FFFFFF')
    
    # 列F～Iのヘッダー背景色
    fg_fill = PatternFill(start_color='4C5D3C', end_color='4C5D3C', fill_type='solid')

    # 定数として各列の列記号を定義
    COLUMN_USER_ID = 'A'       # ユーザーID
    COLUMN_DISCREPANCY = 'B'    # 相違
    COLUMN_STATUS = 'C'         # ステータス
    COLUMN_LOGIN_NAME = 'D'     # ログイン名
    COLUMN_NAME = 'E'           # 氏名
    COLUMN_EMAIL = 'F'          # メールアドレス
    COLUMN_LAST_ACCESS = 'G'    # 最終アクセス日
    COLUMN_DAYS_SINCE = 'H'     # 経過日数
    COLUMN_GROUPS = 'I'         # 所属グループ一覧

    # アクティブと停止中シートのフォーマット
    for sheet in sheets:
        self.logger.info(f"{sheet}シートのフォーマットを設定中...")
        ws = wb[sheet]

        # ヘッダー行（1行目）の各セルに背景色とフォントを設定（A～I列）
        for col in [COLUMN_USER_ID, COLUMN_DISCREPANCY, COLUMN_STATUS, COLUMN_LOGIN_NAME,
                    COLUMN_NAME, COLUMN_EMAIL, COLUMN_LAST_ACCESS, COLUMN_DAYS_SINCE, COLUMN_GROUPS]:
            cell = ws[f'{col}1']
            cell.fill = header_fill
            cell.font = header_font

        # 列幅の設定（ピクセル値を文字数に変換）
        column_widths_px = {
            COLUMN_USER_ID: 180,     # ユーザーID
            COLUMN_DISCREPANCY: 80,   # 相違
            COLUMN_STATUS: 80,       # ステータス
            COLUMN_LOGIN_NAME: 270,  # ログイン名
            COLUMN_NAME: 270,        # 氏名
            COLUMN_EMAIL: 334,       # メールアドレス
            COLUMN_LAST_ACCESS: 160, # 最終アクセス日
            COLUMN_DAYS_SINCE: 60,   # 経過日数
            COLUMN_GROUPS: 1195      # 所属グループ一覧
        }
        for col, px in column_widths_px.items():
            ws.column_dimensions[col].width = px_to_char(px)

        # グループごとの列をJ列以降に設定（幅は15）
        start_col_letter = 'J'
        start_col_num = column_index_from_string(start_col_letter)
        for i, group in enumerate(self.group_names, start=start_col_num):
            col_letter = get_column_letter(i)
            ws.column_dimensions[col_letter].width = 15

        # 列F～I（メールアドレス、最終アクセス日、経過日数、所属グループ一覧）のヘッダーに別背景色を設定
        for col_letter in [COLUMN_EMAIL, COLUMN_LAST_ACCESS, COLUMN_DAYS_SINCE, COLUMN_GROUPS]:
            cell = ws[f'{col_letter}1']
            cell.fill = fg_fill

        # データ行（2行目以降）のセル配置を設定
        exclude_columns = [COLUMN_LOGIN_NAME, COLUMN_NAME, COLUMN_EMAIL]
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                if cell.column_letter in [COLUMN_USER_ID, COLUMN_DISCREPANCY, COLUMN_STATUS,
                                          COLUMN_LAST_ACCESS, COLUMN_DAYS_SINCE]:
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                elif cell.column_letter in exclude_columns:
                    cell.alignment = Alignment(horizontal='left', vertical='center')
                else:
                    cell.alignment = Alignment(horizontal='center', vertical='center')

        # 「Administrators」グループに所属している場合は、氏名（E列）を太字にする
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            group_cell = row[column_index_from_string(COLUMN_GROUPS) - 1]
            if group_cell.value:
                groups = [g.strip() for g in group_cell.value.split(',')]
                if 'Administrators' in groups:
                    name_cell = row[column_index_from_string(COLUMN_NAME) - 1]
                    name_cell.font = Font(bold=True)

        # 所属グループ一覧内に「Administrators」が含まれている場合は除去
        for row in ws.iter_rows(min_row=2, min_col=column_index_from_string(COLUMN_GROUPS),
                                max_col=column_index_from_string(COLUMN_GROUPS), max_row=ws.max_row):
            for cell in row:
                if cell.value and 'Administrators' in cell.value:
                    cell.value = cell.value.replace('Administrators', '').strip()
                    if cell.value.endswith(','):
                        cell.value = cell.value[:-1].strip()

        # 最終アクセス日（G列）と経過日数（H列）は中央寄せ（念のため再設定）
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            last_access_cell = row[column_index_from_string(COLUMN_LAST_ACCESS) - 1]
            days_cell = row[column_index_from_string(COLUMN_DAYS_SINCE) - 1]
            last_access_cell.alignment = Alignment(horizontal='center', vertical='center')
            days_cell.alignment = Alignment(horizontal='center', vertical='center')

        self.logger.info(f"{sheet}シートのフォーマット設定が完了しました。")

    # グループ情報シートのフォーマット
    if 'グループ情報' in wb.sheetnames:
        ws = wb['グループ情報']
        self.logger.info("グループ情報シートのフォーマットを設定中...")
        
        # カラム幅の設定（A～E列）
        column_widths = {
            'A': 15,  # ユーザーID
            'B': 30,  # ログイン名
            'C': 20,  # 氏名
            'D': 35,  # メールアドレス
            'E': 10   # 停止中
        }
        for col, width in column_widths.items():
            ws.column_dimensions[col].width = width
        
        # 背景・フォント設定
        group_title_fill = PatternFill(start_color='5B9BD5', end_color='5B9BD5', fill_type='solid')
        group_title_font = Font(bold=True, color='FFFFFF')
        
        header_fill = PatternFill(start_color='243C5C', end_color='243C5C', fill_type='solid')
        header_font = Font(bold=True, color='FFFFFF')
        
        # 枠線（太線）の設定
        thick_side = Side(border_style='thick', color="000000")
        
        # 非kirin.co.jpドメイン用の背景色
        light_gray_fill = PatternFill(start_color='EEEEEE', end_color='EEEEEE', fill_type='solid')
        
        # シート全体を走査して、各セットごとにフォーマットを適用する
        row = 1
        while row <= ws.max_row:
            cell_val = ws.cell(row=row, column=1).value
            if isinstance(cell_val, str) and cell_val.startswith("グループ:"):
                block_start = row
                # グループ名行の背景設定
                for col in range(1, 6):  # A～E列
                    cell = ws.cell(row=row, column=col)
                    cell.fill = group_title_fill
                    cell.font = group_title_font
                row += 1
                
                # ヘッダー行の背景設定
                if row <= ws.max_row and ws.cell(row=row, column=1).value == "ユーザーID":
                    for col in range(1, 6):  # A～E列
                        cell = ws.cell(row=row, column=col)
                        cell.fill = header_fill
                        cell.font = header_font
                        cell.alignment = Alignment(horizontal='center')  # ヘッダーを中央揃え
                    row += 1
                else:
                    continue
                
                # セットのデータ行の最終行を検出
                data_start = row
                while row <= ws.max_row and ws.cell(row=row, column=1).value not in [None, ""]:
                    # D列（メールアドレス）を右寄せに
                    email_cell = ws.cell(row=row, column=4)
                    email_cell.alignment = Alignment(horizontal='right')
                    
                    # メールアドレスが@kirin.co.jp以外なら背景色を薄いグレーに
                    if email_cell.value and not email_cell.value.lower().endswith('@kirin.co.jp'):
                        email_cell.fill = light_gray_fill
                    
                    # 停止中列を中央揃えに
                    ws.cell(row=row, column=5).alignment = Alignment(horizontal='center')
                    row += 1
                block_end = row - 1
                
                # ブロック全体に太線の枠線を設定
                for r in range(block_start, block_end + 1):
                    for c in range(1, 6):  # A～E列
                        cell = ws.cell(row=r, column=c)
                        new_border = Border(
                            left=thick_side if c == 1 else cell.border.left,
                            right=thick_side if c == 5 else cell.border.right,
                            top=thick_side if r == block_start else cell.border.top,
                            bottom=thick_side if r == block_end else cell.border.bottom
                        )
                        cell.border = new_border
                row += 1
            else:
                row += 1
        
        self.logger.info("グループ情報シートのフォーマット設定が完了しました。")
    
    wb.save(self.output_file)
    self.logger.info(f"Excelファイル '{self.output_file}' のフォーマットを設定しました。")

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
  processor = DataProcessor(all_users, all_groups, client, logger)
  processor.map_users()
  filtered_groups = processor.filter_groups()
  group_names = processor.organize_groups(filtered_groups)
  processor.populate_group_memberships(filtered_groups)
  dataframes = processor.generate_dataframes(group_names)
  
  # group_user_list.yamlの生成を追加
  processor.export_group_user_list(filtered_groups)

  # Excelへのエクスポートとフォーマット
  exporter = ExcelExporter(dataframes, group_names, args.output, logger)
  exporter.prepare_group_data(client)
  exporter.export_to_excel()
  exporter.format_excel()

  logger.info(f"Excelファイル '{args.output}' の作成とフォーマット設定が完了しました。")

if __name__ == "__main__":
  main()