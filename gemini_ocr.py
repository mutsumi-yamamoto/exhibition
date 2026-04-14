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

_PROMPT_SINGLE = """
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

_PROMPT_MULTI = """
以下は同一人物の名刺の表面と裏面の画像です。
両面の情報を統合して、必ずJSON形式のみで返してください。
説明文や前置きは不要です。JSONのみを返してください。

表面に氏名・企業名・役職、裏面にメールアドレス・電話番号・住所などが
記載されている場合があります。両面から読み取れる情報をすべて統合してください。

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


def _prepare_image(image: Image.Image) -> bytes:
    """PIL Image をリサイズ・RGB変換し JPEG bytes を返す。"""
    w, h = image.size
    if max(w, h) > 1024:
        ratio = 1024 / max(w, h)
        image = image.resize((int(w * ratio), int(h * ratio)), Image.LANCZOS)
    if image.mode != "RGB":
        image = image.convert("RGB")
    buf = io.BytesIO()
    image.save(buf, format="JPEG", quality=85)
    return buf.getvalue()


def extract_from_image(images) -> BusinessCard:
    """
    PIL Image（1枚 or 複数枚）を受け取り、Gemini で名刺情報を抽出して BusinessCard を返す。

    Args:
        images: PIL.Image（1枚）または list[PIL.Image]（表裏など複数枚）

    Returns:
        BusinessCard dataclass

    Raises:
        ValueError: APIキー未設定 or JSONパース失敗時
    """
    # 単一画像をリストに正規化
    if isinstance(images, Image.Image):
        images = [images]

    try:
        import streamlit as st
        api_key = st.secrets.get("GEMINI_API_KEY") or os.getenv("GEMINI_API_KEY")
    except Exception:
        api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        raise ValueError("GEMINI_API_KEY が設定されていません（.streamlit/secrets.toml または .env を確認してください）。")

    # 各画像を JPEG bytes に変換
    image_parts = [
        types.Part.from_bytes(data=_prepare_image(img), mime_type="image/jpeg")
        for img in images
    ]

    # 画像枚数に応じてプロンプトを選択
    prompt = _PROMPT_MULTI if len(images) > 1 else _PROMPT_SINGLE

    client = genai.Client(api_key=api_key)

    last_error = None
    for attempt in range(3):
        try:
            response = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=[prompt] + image_parts,
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
