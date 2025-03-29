import json
from typing import List

def parse_dict_string(value_str):
    """
    Python辞書のような文字列を解析し、各項目をname (code)形式の文字列リストに変換する
    
    Args:
        value_str: 辞書リストのような文字列
        
    Returns:
        list: 各辞書を "name (code)" 形式に変換した文字列のリスト
    """
    result = []
    
    # 空の場合は空リストを返す
    if not value_str or not isinstance(value_str, str):
        return result
    
    # 複数の辞書を分割
    # {'code': 'A', 'name': 'B'}, {'code': 'C', 'name': 'D'} --> ['{...}', '{...}']
    items_str = value_str.split("}, {")
    
    for item_str in items_str:
        # 最初と最後の辞書の括弧を修正
        if not item_str.startswith("{"):
            item_str = "{" + item_str
        if not item_str.endswith("}"):
            item_str = item_str + "}"
            
        try:
            # 文字列をPython辞書に変換するために単一引用符を二重引用符に置き換え
            item_str = item_str.replace("'", '"')
            item_dict = json.loads(item_str)
            
            # name (code) 形式の文字列を作成
            if "name" in item_dict and "code" in item_dict:
                result.append(f"{item_dict['name']} ({item_dict['code']})")
            elif "name" in item_dict:
                result.append(item_dict["name"])
            elif "code" in item_dict:
                result.append(item_dict["code"])
        except:
            # 解析に失敗した場合は元の文字列を追加
            result.append(item_str)
    
    return result

def add_field_values_reference(self, ws: Worksheet, row_idx: int, 
                              field_codes: List[str]) -> int:
    """フィールド値の参考一覧をExcelシートに追加する"""
    if not field_codes:
        return row_idx
        
    unique_field_codes = list(set(field_codes))
    for field_code in unique_field_codes:
        values = self.load_field_values_from_tsv(field_code)
        if not values:
            continue
            
        # フィールドタイプ取得（form_fieldsから）
        field_type = ""
        if self.form_fields and field_code in self.form_fields.get('properties', {}):
            field_info = self.form_fields['properties'][field_code]
            field_type = field_info.get('type', '')
            
        # 値の数をログに記録（デバッグ用）
        self.logger.info(f"フィールド '{field_code}' からの値: {len(values)}個")
                
        row_idx += 2
        self.setup_cell(
            ws, row_idx, 1, 
            f"通知先種別：フィールド  フィールドタイプ：{'グループ選択（GROUP_SELECT）' if field_type == 'GROUP_SELECT' else 'ユーザー選択（USER_SELECT）'}",
            fill_color=self.FIELD_HEADER_COLOR
        )
        ws.cell(row=row_idx, column=1).font = Font(bold=True, size=12)
        
        row_idx += 1
        self.setup_cell(
            ws, row_idx, 1, 
            f"フィールド名：{field_code}     ※値は過去データより収集"
        )
        ws.cell(row=row_idx, column=1).font = Font(bold=True)
        
        row_idx += 1

        # ヘッダー追加（フィールドタイプに応じた処理）
        headers = []
        if field_type == 'GROUP_SELECT':
            headers = ["グループ名", "アカウント名", "メールアドレス", "停止中"]
        elif field_type == 'USER_SELECT':
            headers = ["", "アカウント名", "メールアドレス", "停止中"]
        else:
            # その他のタイプでもヘッダーを表示
            headers = ["値"]
            
        # ヘッダー行を追加
        for col_idx, header in enumerate(headers, 1):
            self.setup_cell(ws, row_idx, col_idx, header, is_header=True)
        row_idx += 1

        # 値の表示
        current_row = row_idx
        has_json_values = False  # JSON値があるかのフラグ
        
        for value_idx, value in enumerate(values):
            # 値の種類をログに記録（デバッグ用）
            self.logger.info(f"値 {value_idx+1}: タイプ={type(value)}, 値={value}")
            
            # 1. JSON形式のオブジェクト解析を試みる
            json_objects = []
            try:
                if isinstance(value, str):
                    # 単一のJSONオブジェクト
                    if value.startswith("{") and value.endswith("}"):
                        value_fixed = value.replace("'", '"')
                        obj = json.loads(value_fixed)
                        if 'code' in obj and 'name' in obj:
                            json_objects.append(f"{obj['name']}({obj['code']})")
                    
                    # 複数のJSONオブジェクト
                    elif "[{" in value or "}, {" in value:
                        # 配列形式の修正
                        if value.startswith("[") and not value.startswith("[{"):
                            value = value.replace("[", "[{").replace("]", "}]")
                        
                        parts = value.replace('}, {', '}|{').split('|')
                        for part in parts:
                            part = part.replace("'", '"').strip('[]')
                            if part.startswith("{") and part.endswith("}"):
                                obj = json.loads(part)
                                if 'code' in obj and 'name' in obj:
                                    json_objects.append(f"{obj['name']}({obj['code']})")
            except Exception as e:
                self.logger.warning(f"JSON解析エラー: {e}, 値: {value}")
            
            # 2. JSON解析結果に基づいて表示
            if json_objects:
                has_json_values = True
                for obj_idx, obj_value in enumerate(json_objects):
                    self.setup_cell(ws, current_row, 1, obj_value)
                    current_row += 1
            else:
                # 通常の値として表示（5列使用）
                col = (value_idx % 5) + 1
                self.setup_cell(ws, current_row, col, value)
                if col == 5:  # 5列目まで埋まったら次の行へ
                    current_row += 1
        
        # 最後の通常値の行が途中で終わった場合、次の行へ
        if not has_json_values and values and len(values) % 5 != 0:
            current_row += 1
            
        row_idx = current_row + 1
    
    return row_idx

# 以下の部分を書き換える
# ... existing code ...
            
            # 通常のデータ処理
            if not is_json:
                # Python辞書のような形式を処理
                if isinstance(value, str) and '{' in value and '}' in value:
                    dict_items = parse_dict_string(value)
                    for item in dict_items:
                        col = col_count % 5 + 1
                        cell = ws.cell(row=current_row, column=col)
                        cell.value = item
                        cell.border = thin_border
                        
                        col_count += 1
                        if col_count % 5 == 0:
                            current_row += 1
                else:
                    # 単純な値の処理
                    col = col_count % 5 + 1
                    cell = ws.cell(row=current_row, column=col)
                    cell.value = value
                    cell.border = thin_border
                    
                    col_count += 1
                    if col_count % 5 == 0:
                        current_row += 1
# ... existing code ... 