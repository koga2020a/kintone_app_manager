from openpyxl import Workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont

class ExcelFormatter:
    def __init__(self):
        self.workbook = Workbook()
        self.worksheet = self.workbook.active
    
    def add_data(self, cell_address, text):
        """セルにデータを追加"""
        self.worksheet[cell_address] = text
    
    def bold_text_by_position(self, cell_address, start_pos, length):
        """セル内の特定の位置からの文字をボールドにする"""
        cell = self.worksheet[cell_address]
        
        try:
            original_text = str(cell.value)
            text_length = len(original_text)
            
            # 範囲チェック
            if start_pos < 0 or start_pos >= text_length:
                print(f"開始位置が範囲外です: {start_pos}")
                return
            
            end_pos = min(start_pos + length, text_length)
            
            # テキストブロックを作成
            text_blocks = []
            
            # ボールド前の部分
            if start_pos > 0:
                text_blocks.append(original_text[:start_pos])
            
            # ボールド部分
            bold_text = original_text[start_pos:end_pos]
            bold_font = InlineFont(b=True)
            text_blocks.append(TextBlock(bold_font, bold_text))
            
            # ボールド後の部分
            if end_pos < text_length:
                text_blocks.append(original_text[end_pos:])
            
            cell.value = CellRichText(*text_blocks)
            print(f"セル{cell_address}の位置{start_pos}から{length}文字をボールドにしました")
            
        except Exception as e:
            print(f"エラー: {e}")
    
    def bold_multiple_positions(self, cell_address, positions):
        """セル内の複数位置の文字をボールドにする
        
        Parameters:
        cell_address: セルアドレス（例：'A1'）
        positions: (start_pos, length)のタプルのリスト
        """
        cell = self.worksheet[cell_address]
        
        try:
            original_text = str(cell.value)
            text_length = len(original_text)
            
            # 位置をソート（重複処理を簡単にするため）
            positions = sorted(positions, key=lambda x: x[0])
            
            text_blocks = []
            current_pos = 0
            
            for start_pos, length in positions:
                # 範囲チェック
                if start_pos < 0 or start_pos >= text_length:
                    continue
                
                # 前の部分
                if current_pos < start_pos:
                    text_blocks.append(original_text[current_pos:start_pos])
                
                # ボールド部分
                end_pos = min(start_pos + length, text_length)
                bold_text = original_text[start_pos:end_pos]
                bold_font = InlineFont(b=True)
                text_blocks.append(TextBlock(bold_font, bold_text))
                
                current_pos = end_pos
            
            # 残りの部分
            if current_pos < text_length:
                text_blocks.append(original_text[current_pos:])
            
            cell.value = CellRichText(*text_blocks)
            print(f"セル{cell_address}の複数位置をボールドにしました")
            
        except Exception as e:
            print(f"エラー: {e}")
    
    def save(self, filename):
        """ワークブックを保存"""
        self.workbook.save(filename)

# メイン処理
if __name__ == "__main__":
    formatter = ExcelFormatter()
    formatter.add_data('A1', "Pythonを使ってExcelを操作する方法を学びます。")
    formatter.add_data('A2', "Testデータを確認して処理するTest例です。")
    
    # A1セルの1文字目から6文字をボールドに（"Python"）
    formatter.bold_text_by_position('A1', 0, 6)
    
    # A2セルの複数位置をボールドに
    # "Test"（1文字目から4文字）と"Test"（16文字目から4文字）
    formatter.bold_multiple_positions('A2', [(0, 4), (16, 4)])
    
    # 保存
    formatter.save("sample_excel_cell_richText.py.xlsx")
    print("sample_excel_cell_richText.py.xlsxを保存しました") 