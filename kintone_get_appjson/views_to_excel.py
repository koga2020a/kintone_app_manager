import argparse

def parse_args():
    parser = argparse.ArgumentParser(description='kintoneのビュー設定をExcelに出力する')
    parser.add_argument('app_id', type=int, help='アプリID（必須）')
    parser.add_argument('--subdomain', required=True, help='kintoneのサブドメイン（必須）')
    parser.add_argument('--username', required=True, help='kintoneのユーザー名（必須）')
    parser.add_argument('--password', required=True, help='kintoneのパスワード（必須）')
    parser.add_argument('--api-token', help='kintoneのAPIトークン（オプション）')
    parser.add_argument('--output', required=True, help='出力するExcelファイルのパス（必須）')
    parser.add_argument('--group-master', help='グループマスターファイルのパス（オプション）')
    parser.add_argument('--log-level', choices=['DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL'], default='INFO', help='ログレベル（デフォルト: INFO）')
    return parser.parse_args()

def main():
    args = parse_args()
    # メイン処理の実装 