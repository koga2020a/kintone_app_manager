import os
import sys
import glob
import subprocess
from pathlib import Path

def find_and_open_acl_file(search_char):
    """
    指定された文字で始まるディレクトリを./output配下で探し、
    その中から{search_char}_acl.xlsxファイルを見つけてexplorerで開く
    
    Args:
        search_char (str): 検索する文字
    """
    try:
        # カレントディレクトリの./output パスを取得
        output_dir = Path('./output').resolve()
        
        # outputディレクトリが存在しない場合はエラー
        if not output_dir.exists():
            print(f"エラー: {output_dir} ディレクトリが見つかりません。")
            return
            
        # 検索パターンを作成
        dir_pattern = str(output_dir / f"{search_char}*")
        
        # 該当するディレクトリを検索
        matching_dirs = [d for d in glob.glob(dir_pattern) if os.path.isdir(d)]
        
        if not matching_dirs:
            print(f"エラー: {search_char}* に該当するディレクトリが見つかりません。")
            return
            
        # 各ディレクトリ内でACLファイルを検索
        acl_files = []
        for directory in matching_dirs:
            acl_pattern = os.path.join(directory, f"{search_char}*_acl.xlsx")
            acl_files.extend(glob.glob(acl_pattern))
            
        if not acl_files:
            print(f"エラー: {search_char}*_acl.xlsx に該当するファイルが見つかりません。")
            return
            
        # 見つかったファイルをexplorerで開く
        for file_path in acl_files:
            subprocess.run(['explorer', file_path])
            print(f"開いたファイル: {file_path}")
            
    except Exception as e:
        print(f"エラーが発生しました: {str(e)}")

def main():
    # コマンドライン引数をチェック
    if len(sys.argv) != 2:
        print("使用方法: python script.py <検索文字>")
        return
        
    search_char = sys.argv[1]
    find_and_open_acl_file(search_char)

if __name__ == "__main__":
    main()