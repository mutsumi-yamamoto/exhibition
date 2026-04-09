"""
sheets_writer.py
gspread を使って Google Spreadsheet に名刺データを書き込むモジュール。
"""

import os
from datetime import datetime

import gspread
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv

from gemini_ocr import BusinessCard

load_dotenv()

# Google Sheets & Drive の読み書きスコープ
_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# スプレッドシートのヘッダー定義（列順）
HEADERS = [
    "企業名",               # A: Form/OCR
    "氏名",                 # B: Form/OCR
    "役職",                 # C: Form/OCR
    "メールアドレス",       # D: Form/OCR
    "部署",                 # E: Form/OCR
    "電話番号",             # F: Form/OCR
    "関心事項",             # G: Formのみ
    "取得元",               # H: Form/OCR
    "登録日時",             # I: 自動
    "企業規模（売上高）",   # J: GAS自動
    "企業規模（従業員数）", # K: GAS自動
    "EDINETコード",         # L: GAS自動
    "ランク",               # M: GAS自動
    "ランク判定日時",       # N: GAS自動
    "自社担当者名",         # O: 手入力
    "自社担当者メール",     # P: 手入力
    "下書き作成済",         # Q: GAS自動
    "下書き作成日時",       # R: GAS自動
]


def _get_client() -> gspread.Client:
    """サービスアカウントで認証し gspread クライアントを返す。

    優先順位:
    1. st.secrets["gcp_service_account"]（Streamlit Community Cloud / ローカルの secrets.toml）
    2. SERVICE_ACCOUNT_JSON 環境変数で指定したJSONファイル（ローカル開発）
    """
    try:
        import streamlit as st
        if "gcp_service_account" in st.secrets:
            creds = Credentials.from_service_account_info(
                dict(st.secrets["gcp_service_account"]), scopes=_SCOPES
            )
            return gspread.authorize(creds)
    except Exception:
        pass

    # フォールバック: ローカルのJSONファイル
    json_path = os.getenv("SERVICE_ACCOUNT_JSON", "service_account.json")
    if not os.path.exists(json_path):
        raise FileNotFoundError(
            f"サービスアカウントのJSONファイルが見つかりません: {json_path}\n"
            ".streamlit/secrets.toml に [gcp_service_account] を設定するか、"
            "GCP Console からダウンロードしたJSONをプロジェクトルートに配置してください。"
        )
    creds = Credentials.from_service_account_file(json_path, scopes=_SCOPES)
    return gspread.authorize(creds)


def _get_or_create_sheet(client: gspread.Client) -> gspread.Worksheet:
    """
    スプレッドシートを開き、対象シートを返す。
    シートが存在しない場合は作成してヘッダーを書き込む。
    """
    spreadsheet_id = os.getenv("SPREADSHEET_ID")
    if not spreadsheet_id:
        raise ValueError("SPREADSHEET_ID が .env に設定されていません。")

    sheet_name = os.getenv("SHEET_NAME", "マスター")
    spreadsheet = client.open_by_key(spreadsheet_id)

    try:
        worksheet = spreadsheet.worksheet(sheet_name)
    except gspread.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(title=sheet_name, rows=1000, cols=len(HEADERS))

    # ヘッダーが未設定の場合は書き込む
    existing = worksheet.row_values(1)
    if not existing or existing[0] != HEADERS[0]:
        worksheet.insert_row(HEADERS, index=1)

    return worksheet


def append_business_card(card: BusinessCard, source: str = "名刺") -> int:
    """
    BusinessCard をスプレッドシートに1行追記する。

    Args:
        card: 書き込む名刺データ
        source: 取得元ラベル（"名刺" or "Googleフォーム"）

    Returns:
        書き込んだ行番号（1始まり）
    """
    client = _get_client()
    worksheet = _get_or_create_sheet(client)

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    row = [
        card.company_name,  # A: 企業名
        card.full_name,     # B: 氏名
        card.title,         # C: 役職
        card.email,         # D: メールアドレス
        card.department,    # E: 部署
        card.phone,         # F: 電話番号
        "",                 # G: 関心事項（Formのみ）
        source,             # H: 取得元
        now,                # I: 登録日時
        "",                 # J: 企業規模（売上高）GAS
        "",                 # K: 企業規模（従業員数）GAS
        "",                 # L: EDINETコード GAS
        "",                 # M: ランク GAS
        "",                 # N: ランク判定日時 GAS
        "",                 # O: 自社担当者名 手入力
        "",                 # P: 自社担当者メール 手入力
        "FALSE",            # Q: 下書き作成済 GAS
        "",                 # R: 下書き作成日時 GAS
    ]

    worksheet.append_row(row, value_input_option="USER_ENTERED")

    # 追記後の最終行番号を返す
    return len(worksheet.get_all_values())


def check_duplicate(email: str) -> bool:
    """
    同一メールアドレスがシートに既に存在するか確認する。

    Args:
        email: チェックするメールアドレス

    Returns:
        True: 重複あり / False: 重複なし
    """
    if not email:
        return False

    client = _get_client()
    worksheet = _get_or_create_sheet(client)

    # メールアドレス列（5列目=インデックス4）を取得
    all_values = worksheet.get_all_values()
    email_col_index = HEADERS.index("メールアドレス")

    for row in all_values[1:]:  # ヘッダー行をスキップ
        if len(row) > email_col_index and row[email_col_index].strip() == email.strip():
            return True

    return False
