# kintone_group_cli

kintone のグループを操作するコマンドラインツールです。ユーザーの検索、グループ一覧の表示、
ユーザーの所属グループ設定（追加・削除）ができます。

通常はリポジトリルートの `kintone_runner.py group ...` から呼び出されます。

## 設定

接続情報はリポジトリルートの `.kintone.env`（YAML 形式）を使います。既定でこのファイルを読むので、
通常は `--config` を指定する必要はありません（別の場所にある場合だけ `--config PATH` で指定します）。
runner から実行する場合は自動で渡されます。

```yaml
subdomain: "your-subdomain"   # example.cybozu.com の example 部分
username: "admin@example.com" # 管理者権限を持つユーザーのログイン名
password: "your-password"
```

グループの参照・変更には kintone の**管理者権限**が必要です。

## group_cli.py の使い方

```bash
python group_cli.py [command] [user] [group] [--config PATH] [--silent] [--debug] [--search]
```

| 引数 / オプション | 内容 |
| --- | --- |
| `command` | `list`（グループ一覧）、`set`（所属グループの設定）、または検索キーワード |
| `user` | 対象ユーザーのログイン名（`set` で使用） |
| `group` | 設定するグループ名またはグループコード（`set` で使用） |
| `--config` | 設定ファイルのパス（既定: リポジトリ直下の `.kintone.env`） |
| `--silent` | 詳細なログを表示しない |
| `--debug` | デバッグログを表示する |
| `--search` | 第一引数を検索キーワードとして扱う（`list` や `set` を検索したいとき） |

```bash
# グループ一覧
python group_cli.py list

# ユーザー検索（ログイン名・表示名・メールアドレスの部分一致）
python group_cli.py --search yamada

# ユーザーをグループに設定
python group_cli.py set yamada 営業部

# グループを省略すると、一覧から対話で選択する
python group_cli.py set yamada
```

検索結果が 1 名なら、そのユーザーの所属グループ一覧をそのまま表示します。
複数名の場合は番号を選ぶと所属グループ一覧を表示します（システムグループは除外）。

## runner のコマンドとの対応

```bash
python kintone_runner.py group <サブコマンド> [引数]
```

| runner のサブコマンド | 実行される `group_cli.py` の引数 | 内容 |
| --- | --- | --- |
| `group list` | `list` | グループ一覧を表示する |
| `group search <キーワード>` | `--search <キーワード>` | ユーザーを検索し、所属グループを表示する |
| `group add <ユーザー> <グループ>` | `set <ユーザー> <グループ>` | ユーザーの所属グループを指定して設定する |
| `group remove <ユーザー>` | `set <ユーザー>` | グループを省略した `set` として呼ばれ、対話でグループを選び直す |

`set` は「指定したグループだけに所属させる」動作です。指定しなかったグループからは
ユーザーが外れるため、実行前に内容をよく確認してください。

## 注意事項

- 動的グループのグループコードは指定できません。
- 出力先の Excel ファイルを開いたままだとエラーになります。閉じてから実行してください。
- 認証エラーが出るときは、`.kintone.env` の値と、そのユーザーの権限を確認してください。
