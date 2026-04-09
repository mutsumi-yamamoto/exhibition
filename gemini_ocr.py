"""
gemini_ocr.py
名刺画像をGemini Vision APIに送り、構造化データとして抽出するモジュール。
"""

import os
import json
from dataclasses import dataclass, asdict

import google.generativeai as genai
from PIL import Image
from dotenv import load_dotenv

load_dotenv()

# --------------------------------------------------------------------------- #
# データモデル
# --------------------------------------------------------------------------- #

@dataclass
class BusinessCard:
    """名刺から抽出する項目"""
    company_name: str = ""       # 企業名
    department: str = ""         # 部署名
    title: str = ""              # 役職
    full_name: str = ""          # 氏名
    email: str = ""              # メールアドレス
    phone: str = ""              # 電話番号
    address: str = ""            # 住所
    website: str = ""            # ウェブサイト

    def to_dict(self) -> dict:
        return asdict(self)


# --------------------------------------------------------------------------- #
# Gemini クライアント
# --------------------------------------------------------------------------- #

_PROMPT = """
以下の名刺画像から情報を抽出し、必ずJSON形式のみで返してください。
説明文や前置きは不要です。JSONのみを返してください。

抽出するフィールド:
- company_name: 企業名（法人格含む。例: 株式会社〇〇）
- department: 部署名（なければ空文字）
- title: 役職名（例: 代表取締役、営業部長）
- full_name: 氏名（フルネーム）
- email: メールアドレス
- phone: 電話番号（最初の1件）
- address: 住所（なければ空文字）
- website: ウェブサイトURL（なければ空文字）

出力例:
{
  "company_name": "株式会社サンプル",
  "department": "営業部",
  "title": "部長",
  "full_name": "山田 太郎",
  "email": "yamada@sample.co.jp",
  "phone": "03-1234-5678",
  "address": "東京都渋谷区〇〇1-2-3",
  "website": "https://www.sample.co.jp"
}

読み取れない項目は空文字にしてください。JSONのみ返してください。
"""


def extract_from_image(image: Image.Image) -> BusinessCard:
    """
    PIL Image を受け取り、Gemini で名刺情報を抽出して BusinessCard を返す。

    Args:
        image: アップロードされた名刺画像（PIL.Image）

    Returns:
        BusinessCard dataclass

    Raises:
        ValueError: APIキー未設定 or JSONパース失敗時
    """
    try:
        import streamlit as st
        api_key = st.secrets.get("GEMINI_API_KEY") or os.getenv("GEMINI_API_KEY")
    except Exception:
        api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        raise ValueError("GEMINI_API_KEY が設定されていません（.streamlit/secrets.toml または .env を確認してください）。")

    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-1.5-flash")

    response = model.generate_content(
        [_PROMPT, image],
        generation_config=genai.GenerationConfig(
            response_mime_type="application/json",
            temperature=0.0,
        ),
    )

    raw = response.text.strip()

    # コードブロックが混入した場合の除去
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
        raw = raw.strip()

    try:
        data = json.loads(raw)
    except json.JSONDecodeError as e:
        raise ValueError(f"Gemini のレスポンスをJSONとして解析できませんでした。\n生レスポンス:\n{raw}") from e

    return BusinessCard(
        company_name=data.get("company_name", ""),
        department=data.get("department", ""),
        title=data.get("title", ""),
        full_name=data.get("full_name", ""),
        email=data.get("email", ""),
        phone=data.get("phone", ""),
        address=data.get("address", ""),
        website=data.get("website", ""),
    )
