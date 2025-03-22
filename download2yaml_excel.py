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