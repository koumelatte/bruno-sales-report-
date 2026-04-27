# Bruno カフェ 日次売上報告アプリ

カフェの日々の売上データをWebフォームから入力し、Google スプレッドシートへ自動記録するWebアプリです。

## 概要

手書きや手入力で行っていた日次売上報告をデジタル化し、データ管理・集計を効率化します。  
スタッフが日次業務の締め作業としてブラウザから入力するだけで、Google Sheetsに自動で記録されます。

## 主な機能

- 日次売上データの入力フォーム（天候・気温・売上・客数・廃棄数・試飲数・メモ）
- Google Sheets API による指定セルへの自動バッチ更新
- 日付に対応する行を自動で特定して書き込み

## 技術スタック

| 分類 | 技術 |
|------|------|
| バックエンド | Python / Flask |
| データ連携 | Google Sheets API v4 |
| 認証 | Google サービスアカウント |
| 実行環境 | ローカル（port 5001） |

## セットアップ

### 1. 依存パッケージのインストール

```bash
pip install -r requirements.txt
```

### 2. Google API の認証設定

1. [Google Cloud Console](https://console.cloud.google.com/) でプロジェクトを作成
2. 「APIとサービス」→「ライブラリ」で **Google Sheets API** を有効化
3. サービスアカウントを作成し、JSONキーを `credentials/service_account.json` に配置
4. 対象のスプレッドシートをサービスアカウントのメールアドレスに共有

### 3. 環境変数の設定

```bash
cp .env.example .env
# .env を編集して SPREADSHEET_ID を設定
```

`.env` の例：
```
SPREADSHEET_ID=your_spreadsheet_id_here
CREDENTIALS_PATH=credentials/service_account.json
```

### 4. 起動

```bash
python app.py
```

ブラウザで `http://localhost:5001` にアクセスしてフォームを表示します。

## ファイル構成

```
bruno/
├── app.py               # Flaskメインアプリ
├── requirements.txt
├── credentials/         # APIキー置き場（git管理外）
├── src/
│   ├── sheets_client.py
│   └── config.py
└── templates/           # HTMLフォーム
```

## 作成背景

製造業での生産管理・工程管理の経験をもとに、現場スタッフが使いやすいシンプルな日次報告フローを設計しました。Google Sheetsを活用することで、導入コストを抑えながらデータの一元管理を実現しています。
