"""
gemini_ocr.py
名刺画像をGemini Vision APIに送り、構造化データとして抽出するモジュール。
"""

import io
import os
import json
import time
from dataclasses import dataclass, asdict

from google import genai
from google.genai import types
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
    full_name: str = ""          # 氏名
    title: str = ""              # 役職
    email: str = ""              # メールアドレス
    department: str = ""         # 部署
    phone: str = ""              # 電話番号

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
- full_name: 氏名（フルネーム）
- title: 役職名（例: 代表取締役、営業部長）
- email: メールアドレス
- department: 部署名（なければ空文字）
- phone: 電話番号（最初の1件）

出力例:
{
  "company_name": "株式会社サンプル",
  "full_name": "山田 太郎",
  "title": "部長",
  "email": "yamada@sample.co.jp",
  "department": "営業部",
  "phone": "03-1234-5678"
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

    # PIL Image → JPEG bytes に変換
    buf = io.BytesIO()
    image.save(buf, format="JPEG")
    image_bytes = buf.getvalue()

    client = genai.Client(api_key=api_key)

    last_error = None
    for attempt in range(3):
        try:
            response = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=[
                    _PROMPT,
                    types.Part.from_bytes(data=image_bytes, mime_type="image/jpeg"),
                ],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    temperature=0.0,
                ),
            )
            break
        except Exception as e:
            error_str = str(e)
            # 日次クォータ超過はリトライ不可
            if "PerDay" in error_str or "per_day" in error_str.lower():
                raise ValueError(
                    "本日のAPI使用上限（無料枠20回/日）に達しました。\n"
                    "Google Cloud コンソールで課金を有効にすると上限が解除されます。"
                ) from e
            last_error = e
            if attempt < 2:
                time.sleep(3)
    else:
        raise last_error

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
        full_name=data.get("full_name", ""),
        title=data.get("title", ""),
        email=data.get("email", ""),
        department=data.get("department", ""),
        phone=data.get("phone", ""),
    )
