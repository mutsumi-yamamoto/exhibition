# 展示会クライアント登録システム

展示会で取得した名刺をOCRで読み取り、Google Sheetsに一括登録するStreamlitアプリです。

## 機能

| タブ | 概要 |
|---|---|
| カメラで撮影 | スマホ/PCのカメラで名刺を撮影→OCR→登録 |
| ファイルから読取 | JPG/PNG/WEBP/PDFをアップロード→OCR→登録（複数ファイル対応） |
| Driveから読取 | Google Driveのフォルダから名刺画像を一括取得→OCR→登録 |
| 手動入力 | OCRを使わず直接フォームに入力→登録 |

- Gemini 2.5 Flash による名刺OCR（表裏PDF統合対応）
- Google Sheetsへの自動書き込み・重複チェック
- Google Driveへの名刺画像アップロード・共有リンク生成
- フォルダ名からコース自動判定（`01`→事業相談、`02`→AI研修、`03`→システムリプレイス）

## ファイル構成

```
├── app.py              # Streamlitメインアプリ（UI・画面遷移）
├── gemini_ocr.py       # Gemini Vision APIで名刺→BusinessCard dataclass
├── sheets_writer.py    # Google Sheets書き込み・Drive操作（アップロード/ダウンロード）
├── requirements.txt    # Pythonパッケージ依存
├── .env.example        # 環境変数テンプレート
└── .gitignore
```

## セットアップ

### 1. GCPサービスアカウント作成

1. [GCP Console](https://console.cloud.google.com/) でプロジェクトを作成
2. 以下のAPIを有効化:
   - Google Sheets API
   - Google Drive API
3. **サービスアカウント**を作成し、JSONキーをダウンロード
4. ダウンロードしたJSONを `service_account.json` としてプロジェクトルートに配置

### 2. Google Sheets 準備

1. スプレッドシートを作成
2. サービスアカウントのメールアドレス（`xxx@xxx.iam.gserviceaccount.com`）を**編集者**として共有
3. スプレッドシートのURLから **SPREADSHEET_ID** を取得（`/d/` と `/edit` の間の文字列）

### 3. Google Drive 準備

2つのDriveフォルダを用意します:

- **読み取り元フォルダ**: 名刺画像を格納するフォルダ（コース別サブフォルダを含む）
- **保管先フォルダ**: OCR後の名刺画像をアップロードする先

読み取り元フォルダの構成:
```
読み取り元フォルダ/
├── 01.ヤフー元CEO小澤の事業相談/
│   ├── 名刺001.jpg
│   └── 名刺002.pdf
├── 02.AIエージェント研修/
└── 03.システムリプレイス/
```

各フォルダをサービスアカウントのメールアドレスと共有してください。

フォルダIDはURLの `/folders/` の後ろの文字列です:
```
https://drive.google.com/drive/folders/【ここがフォルダID】
```

### 4. Gemini API キー取得

[Google AI Studio](https://aistudio.google.com/) でAPIキーを取得してください。

### 5. 環境変数の設定

#### ローカル開発の場合

`.env.example` をコピーして `.env` を作成し、値を埋めてください:

```bash
cp .env.example .env
```

```env
GEMINI_API_KEY=あなたのGemini APIキー
SPREADSHEET_ID=スプレッドシートのID
SHEET_NAME=マスター
DRIVE_SOURCE_FOLDER_ID=読み取り元フォルダのID
DRIVE_UPLOAD_FOLDER_ID=保管先フォルダのID
```

#### Streamlit Community Cloud の場合

アプリの **Settings > Secrets** に以下を設定:

```toml
GEMINI_API_KEY = "あなたのGemini APIキー"
SPREADSHEET_ID = "スプレッドシートのID"
SHEET_NAME = "マスター"
DRIVE_SOURCE_FOLDER_ID = "読み取り元フォルダのID"
DRIVE_UPLOAD_FOLDER_ID = "保管先フォルダのID"

[gcp_service_account]
type = "service_account"
project_id = "your-project-id"
private_key_id = "key-id"
private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
client_email = "xxx@xxx.iam.gserviceaccount.com"
client_id = "123456789"
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "your-cert-url"
```

`[gcp_service_account]` にはダウンロードしたサービスアカウントJSONの中身をTOML形式で記載します。

### 6. ローカル起動

```bash
pip install -r requirements.txt
streamlit run app.py
```

## スプレッドシートのカラム構成

| 列 | 項目 | 入力元 |
|---|---|---|
| A | 企業名 | OCR/手動 |
| B | 氏名 | OCR/手動 |
| C | 役職 | OCR/手動 |
| D | メールアドレス | OCR/手動 |
| E | 部署 | OCR/手動 |
| F | 電話番号 | OCR/手動 |
| G | 関心事項（コース） | フォーム選択 |
| H | 取得元 | 自動（"名刺OCR" or "手動入力"） |
| I | 登録日時 | 自動 |
| J | 企業規模（売上高） | GAS自動 |
| K | 企業規模（従業員数） | GAS自動 |
| L | EDINETコード | GAS自動 |
| M | ランク | GAS自動 |
| N | ランク判定日時 | GAS自動 |
| O | 自社担当者名 | 手入力 |
| P | 自社担当者メール | 手入力 |
| Q | 下書き作成済 | GAS自動 |
| R | 下書き作成日時 | GAS自動 |
| S | メール送信済 | GAS自動 |
| T | メール送信日時 | GAS自動 |
| U | 名刺画像URL | 自動（Drive） |

## 技術スタック

- **Python** / Streamlit
- **Gemini 2.5 Flash** — 名刺OCR
- **gspread** — Google Sheets読み書き
- **Google Drive API v3** — 画像アップロード/ダウンロード
- **PyMuPDF** — PDF→画像変換
