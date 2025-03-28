import json

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