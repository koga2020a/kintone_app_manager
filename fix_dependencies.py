#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
NumPyとPandasの互換性問題を解決するスクリプト

このスクリプトは、NumPyとPandasの互換性問題を解決するために
必要なパッケージを適切なバージョンでインストールし直します。
"""

import subprocess
import sys
import os
import logging
from pathlib import Path
from datetime import datetime

# ログ設定
def setup_logging():
    """ロギングの設定"""
    log_dir = Path(__file__).resolve().parent / "logs"
    log_dir.mkdir(exist_ok=True)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = log_dir / f"dependency_fix_{timestamp}.log"
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()
        ]
    )
    
    return logging.getLogger("dependency_fix")

def run_command(command, logger):
    """
    コマンドを実行し、結果をログに出力する
    """
    logger.info(f"コマンドを実行: {' '.join(command)}")
    try:
        result = subprocess.run(command, check=True, capture_output=True, text=True)
        logger.info(f"コマンド実行成功: {result.stdout}")
        return True
    except subprocess.CalledProcessError as e:
        logger.error(f"コマンド実行エラー: {e}")
        logger.error(f"標準出力: {e.stdout}")
        logger.error(f"標準エラー: {e.stderr}")
        return False

def fix_dependencies(logger):
    """
    NumPyとPandasの互換性問題を解決する
    """
    # 現在のパッケージバージョンを確認
    logger.info("現在インストールされているパッケージのバージョンを確認")
    run_command([sys.executable, "-m", "pip", "list"], logger)
    
    # まずnumpyを最新バージョンに更新
    logger.info("NumPyをアンインストール中...")
    run_command([sys.executable, "-m", "pip", "uninstall", "-y", "numpy"], logger)
    
    # pandasをアンインストール
    logger.info("Pandasをアンインストール中...")
    run_command([sys.executable, "-m", "pip", "uninstall", "-y", "pandas"], logger)
    
    # openpyxlをアンインストール (Pandasに関連するため)
    logger.info("openpyxlをアンインストール中...")
    run_command([sys.executable, "-m", "pip", "uninstall", "-y", "openpyxl"], logger)
    
    # pip自体を更新
    logger.info("pipを最新バージョンに更新中...")
    run_command([sys.executable, "-m", "pip", "install", "--upgrade", "pip"], logger)
    
    # numpy, pandas, openpyxlを互換性のあるバージョンでインストール
    logger.info("NumPyを互換性のあるバージョンでインストール中...")
    numpy_result = run_command([sys.executable, "-m", "pip", "install", "numpy==1.22.4"], logger)
    
    if numpy_result:
        logger.info("Pandasをインストール中...")
        pandas_result = run_command([sys.executable, "-m", "pip", "install", "pandas==1.4.3"], logger)
        
        if pandas_result:
            logger.info("openpyxlをインストール中...")
            openpyxl_result = run_command([sys.executable, "-m", "pip", "install", "openpyxl==3.0.10"], logger)
            
            if openpyxl_result:
                logger.info("依存関係の修正が完了しました")
                return True
    
    logger.error("依存関係の修正に失敗しました")
    return False

def main():
    """メイン関数"""
    logger = setup_logging()
    logger.info("NumPyとPandasの互換性問題の修正を開始します")
    
    if fix_dependencies(logger):
        logger.info("修正が完了しました。kintone_runner.pyを再実行してください。")
        print("\n修正が完了しました。以下のコマンドでkintone_runnerを再実行してください：")
        print("python kintone_runner.py users")
    else:
        logger.error("修正に失敗しました。ログを確認してください。")
        print("\n修正に失敗しました。手動で以下のコマンドを実行してください：")
        print("pip uninstall -y numpy pandas openpyxl")
        print("pip install numpy==1.22.4 pandas==1.4.3 openpyxl==3.0.10")
    
    logger.info("スクリプトを終了します")

if __name__ == "__main__":
    main() 