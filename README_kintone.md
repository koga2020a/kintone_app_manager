# kintoneの認証とREST API

## 認証の種類

### APIキー認証
- **特徴**: APIキーを使用して認証を行う。簡単に設定できるが、キーの管理が重要。
- **呼び出し方**: HTTPヘッダーに`X-Cybozu-API-Token`を設定。

### ユーザー名・パスワード認証
- **特徴**: ユーザー名とパスワードをBase64エンコードして使用。セキュリティ上の理由から、推奨されない場合もある。
- **呼び出し方**: HTTPヘッダーに`Authorization: Basic {base64_encoded_credentials}`を設定。

## その他の注意点
- APIキーやユーザー名・パスワードの管理は厳重に行うこと。
- HTTPSを使用して通信を暗号化することが推奨される。
