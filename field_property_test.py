import json
import os
import csv
import sys
from typing import Union


class PropertyInfo:
    def __init__(self, key, code, is_subtable=False, subtable_key=None):
        self.key = key
        self.code = code
        self.is_subtable = is_subtable
        self.subtable_key = subtable_key

    def __repr__(self):
        return f"PropertyInfo(key='{self.key}', code='{self.code}', is_subtable={self.is_subtable}, subtable_key='{self.subtable_key}')"


class PropertyMapper:
    def __init__(self, properties: dict):
        self.key_to_info = {}
        self.code_to_info = {}
        self._parse_properties(properties)

    @classmethod
    def from_json_file(cls, path: str):
        if not os.path.isfile(path):
            raise FileNotFoundError(f"ファイルが存在しません: {path}")

        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except json.JSONDecodeError as e:
            raise ValueError(f"JSONの読み込みに失敗しました: {e}")

        if "properties" not in data or not isinstance(data["properties"], dict):
            raise ValueError("JSONの形式が不正です: 'properties' フィールドが見つかりません")

        return cls(data["properties"])

    def _parse_properties(self, properties: dict):
        for key, value in properties.items():
            code = value.get("code")
            prop_type = value.get("type")

            if prop_type == "SUBTABLE":
                fields = value.get("fields", {})
                for sub_key, sub_value in fields.items():
                    sub_code = sub_value.get("code")
                    info = PropertyInfo(
                        key=sub_key,
                        code=sub_code,
                        is_subtable=True,
                        subtable_key=key
                    )
                    self.key_to_info[sub_key] = info
                    self.code_to_info[sub_code] = info

                table_info = PropertyInfo(
                    key=key,
                    code=code,
                    is_subtable=False,
                    subtable_key=None
                )
                self.key_to_info[key] = table_info
                self.code_to_info[code] = table_info
            else:
                info = PropertyInfo(
                    key=key,
                    code=code,
                    is_subtable=False,
                    subtable_key=None
                )
                self.key_to_info[key] = info
                self.code_to_info[code] = info

    def get_by_key(self, key):
        return self.key_to_info.get(key)

    def get_by_code(self, code):
        return self.code_to_info.get(code)

    def get_display_key_by_code(self, code: str) -> Union[str, None]:
        info = self.get_by_code(code)
        if not info:
            return None
        if info.is_subtable:
            subtable_info = self.get_by_key(info.subtable_key)
            return "{}[{}]".format(subtable_info.key, info.key)
        return info.key

    def get_display_code_by_code(self, code: str) -> Union[str, None]:
        info = self.get_by_code(code)
        if not info:
            return None
        if info.is_subtable:
            subtable_info = self.get_by_key(info.subtable_key)
            return "{}[{}]".format(subtable_info.code, info.code)
        return info.code

    def export_debug_info(self, filename: str):
        try:
            with open(filename, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(["display_key", "display_code", "is_subtable", "subtable_key"])
                for info in self.code_to_info.values():
                    display_key = self.get_display_key_by_code(info.code)
                    display_code = self.get_display_code_by_code(info.code)
                    writer.writerow([
                        display_key,
                        display_code,
                        str(info.is_subtable),
                        info.subtable_key or ""
                    ])
            print("[✓] フィールド情報を '{}' に出力しました。".format(filename))
        except Exception as e:
            print("[✗] 出力失敗: {}".format(e))


# --- メイン処理 ---
if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("使い方: python property_mapper_tool.py <input_json_path> <output_csv_path>")
        sys.exit(1)

    json_path = sys.argv[1]
    output_path = sys.argv[2]

    try:
        mapper = PropertyMapper.from_json_file(json_path)
        mapper.export_debug_info(output_path)
    except Exception as e:
        print("[エラー] {}".format(e))
        sys.exit(1)
