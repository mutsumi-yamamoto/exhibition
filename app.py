"""
app.py
名刺OCR / 手動入力 → 確認・修正 → Google Sheets書き込み を行う Streamlit アプリ。

起動方法:
    streamlit run app.py
"""

import io
import time
from typing import Optional
from PIL import Image
import streamlit as st

from gemini_ocr import extract_from_image, BusinessCard
from sheets_writer import append_business_card, check_duplicate

# --------------------------------------------------------------------------- #
# ページ設定
# --------------------------------------------------------------------------- #

st.set_page_config(
    page_title="クライアント登録",
    page_icon="📇",
    layout="centered",
)

st.title("📇 クライアント登録システム")

# --------------------------------------------------------------------------- #
# セッション状態の初期化
# --------------------------------------------------------------------------- #

if "card" not in st.session_state:
    st.session_state.card: Optional[BusinessCard] = None
if "ocr_done" not in st.session_state:
    st.session_state.ocr_done = False
if "submitted" not in st.session_state:
    st.session_state.submitted = False
if "last_photo_id" not in st.session_state:
    st.session_state.last_photo_id = None
if "last_upload_id" not in st.session_state:
    st.session_state.last_upload_id = None
if "dup_cache" not in st.session_state:
    st.session_state.dup_cache: dict = {}
if "form_key" not in st.session_state:
    st.session_state.form_key = 0


def _check_duplicate_cached(email: str) -> bool:
    """重複チェック結果をセッション内でキャッシュし、API呼び出しを最小化する。"""
    if email not in st.session_state.dup_cache:
        try:
            st.session_state.dup_cache[email] = check_duplicate(email)
        except Exception:
            return False
    return st.session_state.dup_cache[email]


# フォームのキー（登録完了後にインクリメントしてフィールドをリセット）
fk = st.session_state.form_key

# --------------------------------------------------------------------------- #
# メインタブ
# --------------------------------------------------------------------------- #

tab_ocr, tab_manual = st.tabs(["📷 名刺OCR", "✏️ 手動入力"])

# =========================================================================== #
# タブ1: 名刺OCR
# =========================================================================== #

with tab_ocr:

    # ----------------------------------------------------------------------- #
    # Step 1: 画像取得（カメラ or アップロード）
    # ----------------------------------------------------------------------- #

    st.markdown("### Step 1 — 名刺画像を取得")

    cam_tab, upload_tab = st.tabs(["📷 カメラで撮影", "📁 ファイルをアップロード"])

    image = None
    image_caption = ""
    new_image = False

    with cam_tab:
        camera_photo = st.camera_input("名刺をカメラで撮影してください", key=f"camera_{fk}", facing_mode="environment")
        if camera_photo is not None:
            image = Image.open(camera_photo)
            image_caption = "撮影した名刺"
            if st.session_state.last_photo_id != id(camera_photo):
                st.session_state.last_photo_id = id(camera_photo)
                st.session_state.last_upload_id = None
                new_image = True

    with upload_tab:
        uploaded_file = st.file_uploader(
            "名刺の画像ファイルを選択してください（JPG / PNG / WEBP）",
            type=["jpg", "jpeg", "png", "webp"],
            help="鮮明に撮影された名刺画像ほど精度が上がります。",
            key=f"uploader_{fk}",
        )
        if uploaded_file is not None:
            image = Image.open(io.BytesIO(uploaded_file.read()))
            image_caption = uploaded_file.name
            upload_key = f"{uploaded_file.name}_{uploaded_file.size}"
            if st.session_state.last_upload_id != upload_key:
                st.session_state.last_upload_id = upload_key
                st.session_state.last_photo_id = None
                new_image = True

    if image is not None:
        col_img, col_info = st.columns([1, 1])
        with col_img:
            st.image(image, caption=image_caption, use_container_width=True)
        with col_info:
            st.info(f"**サイズ**: {image.size[0]} × {image.size[1]} px")

        # ------------------------------------------------------------------- #
        # Step 2: 新しい画像の場合のみ自動OCR
        # ------------------------------------------------------------------- #

        if new_image:
            st.session_state.ocr_done = False
            st.session_state.submitted = False
            with st.spinner("Gemini APIで名刺を解析中..."):
                try:
                    card = extract_from_image(image)
                    st.session_state.card = card
                    st.session_state.ocr_done = True
                    st.success("抽出が完了しました。内容を確認・修正してください。")
                except ValueError as e:
                    st.error(f"エラー: {e}")
                except Exception as e:
                    st.error(f"予期しないエラーが発生しました: {e}")

    # ----------------------------------------------------------------------- #
    # Step 3: 確認・修正フォーム
    # ----------------------------------------------------------------------- #

    if st.session_state.ocr_done and st.session_state.card is not None and not st.session_state.submitted:
        card = st.session_state.card

        st.markdown("### Step 3 — 内容を確認・修正")
        st.caption("OCR結果を確認し、誤りがあれば修正してから登録してください。")

        with st.form(f"ocr_confirm_form_{fk}"):
            col1, col2 = st.columns(2)

            with col1:
                company_name = st.text_input("企業名 *",         value=card.company_name)
                full_name    = st.text_input("氏名 *",           value=card.full_name)
                title        = st.text_input("役職 *",           value=card.title)

            with col2:
                email        = st.text_input("メールアドレス *", value=card.email)
                department   = st.text_input("部署",             value=card.department)
                phone        = st.text_input("電話番号",         value=card.phone)

            if card.email:
                try:
                    if _check_duplicate_cached(card.email):
                        st.warning(
                            f"⚠️ このメールアドレス（{card.email}）は既に登録されています。"
                            "重複登録に注意してください。"
                        )
                except Exception:
                    pass

            submitted = st.form_submit_button("✅ Google Sheetsに登録する", type="primary", use_container_width=True)

            if submitted:
                missing = []
                if not company_name.strip():
                    missing.append("企業名")
                if not full_name.strip():
                    missing.append("氏名")
                if not email.strip():
                    missing.append("メールアドレス")

                if missing:
                    st.error(f"必須項目が未入力です: {', '.join(missing)}")
                else:
                    confirmed_card = BusinessCard(
                        company_name=company_name.strip(),
                        department=department.strip(),
                        title=title.strip(),
                        full_name=full_name.strip(),
                        email=email.strip(),
                        phone=phone.strip(),
                    )
                    with st.spinner("Google Sheetsに書き込み中..."):
                        try:
                            row_num = append_business_card(confirmed_card, source="名刺OCR")
                            st.session_state.submitted = True
                            st.rerun()
                        except FileNotFoundError as e:
                            st.error(f"認証エラー: {e}")
                        except ValueError as e:
                            st.error(f"設定エラー: {e}")
                        except Exception as e:
                            st.error(f"書き込みエラー: {e}")

# =========================================================================== #
# タブ2: 手動入力
# =========================================================================== #

with tab_manual:

    st.markdown("### 情報を手動で入力")
    st.caption("名刺がない場合や、直接入力したい場合にご利用ください。")

    with st.form(f"manual_form_{fk}"):
        col1, col2 = st.columns(2)

        with col1:
            m_company_name = st.text_input("企業名 *",         key=f"m_company_{fk}")
            m_full_name    = st.text_input("氏名 *",           key=f"m_fullname_{fk}")
            m_title        = st.text_input("役職 *",           key=f"m_title_{fk}")

        with col2:
            m_email        = st.text_input("メールアドレス *", key=f"m_email_{fk}")
            m_department   = st.text_input("部署",             key=f"m_dept_{fk}")
            m_phone        = st.text_input("電話番号",         key=f"m_phone_{fk}")

        submitted_manual = st.form_submit_button("✅ Google Sheetsに登録する", type="primary", use_container_width=True)

        if submitted_manual:
            if m_email.strip():
                try:
                    if _check_duplicate_cached(m_email.strip()):
                        st.warning(
                            f"⚠️ このメールアドレス（{m_email.strip()}）は既に登録されています。"
                            "重複登録に注意してください。"
                        )
                except Exception:
                    pass

            missing = []
            if not m_company_name.strip():
                missing.append("企業名")
            if not m_full_name.strip():
                missing.append("氏名")
            if not m_email.strip():
                missing.append("メールアドレス")

            if missing:
                st.error(f"必須項目が未入力です: {', '.join(missing)}")
            else:
                manual_card = BusinessCard(
                    company_name=m_company_name.strip(),
                    full_name=m_full_name.strip(),
                    title=m_title.strip(),
                    email=m_email.strip(),
                    department=m_department.strip(),
                    phone=m_phone.strip(),
                )
                with st.spinner("Google Sheetsに書き込み中..."):
                    try:
                        row_num = append_business_card(manual_card, source="手動入力")
                        st.session_state.submitted = True
                        st.rerun()
                    except FileNotFoundError as e:
                        st.error(f"認証エラー: {e}")
                    except ValueError as e:
                        st.error(f"設定エラー: {e}")
                    except Exception as e:
                        st.error(f"書き込みエラー: {e}")

# --------------------------------------------------------------------------- #
# 登録完了（OCR・手動共通）
# --------------------------------------------------------------------------- #

if st.session_state.submitted:
    st.success("✅ 登録が完了しました。次の読み取りに移ります...")
    time.sleep(2)
    st.session_state.card = None
    st.session_state.ocr_done = False
    st.session_state.submitted = False
    st.session_state.last_photo_id = None
    st.session_state.last_upload_id = None
    st.session_state.dup_cache = {}
    st.session_state.form_key += 1  # フォームキーを更新してフィールドをリセット
    st.rerun()
