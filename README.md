# Bruno - Google Sheets Data Loader

PythonでGoogle Sheetsのデータを読み込むプロジェクト。

## セットアップ

### 1. 依存パッケージのインストール

```bash
pip install -r requirements.txt
```

### 2. Google API の認証設定

#### Service Account の作成

1. [Google Cloud Console](https://console.cloud.google.com/) でプロジェクトを作成
2. 「APIとサービス」→「ライブラリ」で **Google Sheets API** と **Google Drive API** を有効化
3. 「APIとサービス」→「認証情報」→「サービスアカウントを作成」
4. サービスアカウントのJSONキーをダウンロードし `credentials/service_account.json` に配置
5. 対象のGoogle Sheetをそのサービスアカウントのメールアドレスに共有

#### 環境変数の設定

```bash
cp .env.example .env
# .env を編集して SPREADSHEET_ID を設定
```

`.env` の例:
```
SPREADSHEET_ID=your_spreadsheet_id_here
CREDENTIALS_PATH=credentials/service_account.json
```

Spreadsheet IDはスプレッドシートURLの以下の部分:
```
https://docs.google.com/spreadsheets/d/<SPREADSHEET_ID>/edit
```

### 3. 動作確認

```python
from src.sheets_client import SheetsClient

client = SheetsClient()
df = client.read_sheet("Sheet1")
print(df.head())
```

## ファイル構成

```
bruno/
├── README.md
├── .env.example
├── .gitignore
├── requirements.txt
├── credentials/          # APIキー置き場（gitで管理しない）
├── src/
│   ├── sheets_client.py  # Google Sheets接続クライアント
│   └── config.py         # 設定管理
├── notebooks/
│   └── example.ipynb     # 使い方サンプル
└── tests/
    └── test_sheets_client.py
```

## 主な使い方

```python
from src.sheets_client import SheetsClient

client = SheetsClient()

# シート全体をDataFrameとして取得
df = client.read_sheet("Sheet1")

# 範囲を指定して取得
df = client.read_range("Sheet1", "A1:D100")

# 複数シートを取得
sheets = client.list_sheets()
```

## 依存ライブラリ

| ライブラリ | 用途 |
|-----------|------|
| `google-auth` | Google認証 |
| `google-api-python-client` | Google API クライアント |
| `gspread` | Google Sheets 操作 |
| `pandas` | データ処理 |
| `python-dotenv` | 環境変数管理 |
