import sys

class ExcelFormatter:
    # ... existing code ...

    def set_by_out02_tsv(self, tsv_filename):
        """構造化されたTSVからセルを設置（フィールド名をネストレベルに応じて右にずらす）"""
        # ... existing code ...

        for i, row in enumerate(rows):
            # ... existing code ...

            if len(row) > 4:
                field_type = row[4]
                field_type_ja = {
                    # ... existing code ...
                }.get(field_type, field_type)
                set_val_font(self.ws[f'BB{i+3}'], field_type_ja)

                # ドロップダウンの選択肢をBC列に表示
                if field_type == 'DROP_DOWN' and len(row) > 10:
                    options_str = row[10]
                    options = []
                    try:
                        items = options_str.split(',')
                        for item in items:
                            if ': {' in item:
                                option = item.split(': {')[0].strip()
                                if option not in ['options', 'index', 'defaultValue'] and not option.startswith('"'):
                                    options.append(option)
                        if options:
                            set_val_font(self.ws[f'BC{i+3}'], '選択肢: ' + ', '.join(options))
                    except Exception as e:
                        print(f"選択肢の解析エラー: {e}")

            # ... existing code ...

    # ... existing code ... 

def run_download2yaml_excel(app_id: str = None, output_dir: str = None) -> bool:
    """
    download2yaml_excel.pyのメイン処理を実行する関数
    
    Args:
        app_id (str, optional): アプリID
        output_dir (str, optional): 出力ディレクトリ
        
    Returns:
        bool: 処理が成功したかどうか
    """
    try:
        # 引数を設定
        sys.argv = ['download2yaml_excel.py']
        if app_id:
            sys.argv.extend(['--id', app_id])
        if output_dir:
            sys.argv.extend(['--output-dir', output_dir])
            
        # main関数を実行
        main()
        return True
    except Exception as e:
        print(f"エラーが発生しました: {e}")
        return False

def main():
    """メイン関数"""
    # ... existing code ... 