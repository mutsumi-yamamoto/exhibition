"""
app.py
名刺OCR / 手動入力 → 確認・修正 → Google Sheets書き込み を行う Streamlit アプリ。

起動方法:
    streamlit run app.py
"""

import io
import os
import time
from typing import Optional
from PIL import Image
import fitz  # PyMuPDF
import streamlit as st

from gemini_ocr import extract_from_image, BusinessCard
from sheets_writer import (
    append_business_card,
    check_duplicate,
    upload_to_drive,
    list_drive_subfolders,
    list_drive_images,
    download_drive_file,
)

# コース選択肢
COURSE_OPTIONS = [
    "01. ヤフー元CEO小澤の事業相談",
    "02. AIエージェント研修",
    "03. システムリプレイス",
]

# フォルダ名先頭2文字 → コースのマッピング
_FOLDER_COURSE_MAP = {
    "01": COURSE_OPTIONS[0],
    "02": COURSE_OPTIONS[1],
    "03": COURSE_OPTIONS[2],
}


def _folder_to_course(folder_name: str) -> Optional[str]:
    """フォルダ名の先頭2文字からコースを自動判定する。"""
    prefix = folder_name[:2]
    return _FOLDER_COURSE_MAP.get(prefix)


def _pdf_to_images(pdf_bytes: bytes) -> list[Image.Image]:
    """PDFの全ページを PIL Image のリストに変換する。"""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    images = []
    for page in doc:
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        images.append(img)
    doc.close()
    return images


def _to_jpeg_bytes(img: Image.Image) -> bytes:
    """PIL Image を JPEG バイト列に変換（RGBA/Pモードは RGB に変換）。"""
    if img.mode != "RGB":
        img = img.convert("RGB")
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=90)
    return buf.getvalue()

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
if "image_bytes" not in st.session_state:
    st.session_state.image_bytes: Optional[bytes] = None
if "multi_images" not in st.session_state:
    st.session_state.multi_images: list = []
if "multi_idx" not in st.session_state:
    st.session_state.multi_idx = 0
if "multi_upload_key" not in st.session_state:
    st.session_state.multi_upload_key = None
if "drive_subfolders" not in st.session_state:
    st.session_state.drive_subfolders: list = []
if "drive_course" not in st.session_state:
    st.session_state.drive_course: str | None = None


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
# メインタブ（フラット構造）
# --------------------------------------------------------------------------- #

tab_camera, tab_upload, tab_drive, tab_manual = st.tabs(
    ["📷 カメラで撮影", "📁 ファイルから読取", "☁️ Driveから読取", "✏️ 手動入力"]
)

image = None
image_caption = ""
new_image = False


def _run_ocr(images) -> None:
    """画像（1枚 or 複数枚）に対してOCRを実行しセッションに結果を保存。"""
    st.session_state.ocr_done = False
    st.session_state.submitted = False
    with st.spinner("Gemini APIで名刺を解析中..."):
        try:
            card = extract_from_image(images)
            st.session_state.card = card
            st.session_state.ocr_done = True
            st.success("抽出が完了しました。内容を確認・修正してください。")
        except ValueError as e:
            st.error(f"エラー: {e}")
        except Exception as e:
            st.error(f"予期しないエラーが発生しました: {e}")


# =========================================================================== #
# タブ1: カメラで撮影
# =========================================================================== #

with tab_camera:
    st.caption("外カメラを使う場合は、カメラ画面内の切替ボタン（↺）をタップしてください。")
    camera_photo = st.camera_input("名刺をカメラで撮影してください", key=f"camera_{fk}")
    if camera_photo is not None:
        image = Image.open(camera_photo)
        image_caption = "撮影した名刺"
        if st.session_state.last_photo_id != id(camera_photo):
            st.session_state.last_photo_id = id(camera_photo)
            st.session_state.last_upload_id = None
            new_image = True
            # Drive アップロード用にバイト列を保存
            camera_photo.seek(0)
            st.session_state.image_bytes = camera_photo.read()

        col_img, col_info = st.columns([1, 1])
        with col_img:
            st.image(image, caption=image_caption, use_container_width=True)
        with col_info:
            st.info(f"**サイズ**: {image.size[0]} × {image.size[1]} px")

        if new_image:
            _run_ocr(image)

# =========================================================================== #
# タブ2: ファイルから読取
# =========================================================================== #

with tab_upload:
    st.caption("複数ファイルを同時に選択できます。1枚ずつ確認・登録します。")
    uploaded_files = st.file_uploader(
        "名刺の画像ファイルを選択してください（JPG / PNG / WEBP / PDF）",
        type=["jpg", "jpeg", "png", "webp", "pdf"],
        accept_multiple_files=True,
        help="鮮明に撮影された名刺画像ほど精度が上がります。PDFは1ページ目を読み取ります。",
        key=f"uploader_{fk}",
    )
    if uploaded_files:
        # 新しいファイル群かどうかを判定
        upload_key = "_".join(f"{f.name}_{f.size}" for f in uploaded_files)
        if st.session_state.multi_upload_key != upload_key:
            st.session_state.multi_upload_key = upload_key
            st.session_state.last_photo_id = None
            # 全ファイルを画像に変換してキューに格納
            queue = []
            for f in uploaded_files:
                raw = f.read()
                if f.name.lower().endswith(".pdf"):
                    imgs = _pdf_to_images(raw)
                else:
                    imgs = [Image.open(io.BytesIO(raw))]
                queue.append({
                    "images": imgs,              # OCR用（表裏など全ページ）
                    "image": imgs[0],            # プレビュー用（1ページ目）
                    "image_bytes": _to_jpeg_bytes(imgs[0]),  # Drive用
                    "filename": f.name,
                })
            st.session_state.multi_images = queue
            st.session_state.multi_idx = 0
            st.session_state.ocr_done = False
            st.session_state.card = None

        # 現在のキューからアイテムを表示
        queue = st.session_state.multi_images
        idx = st.session_state.multi_idx

        if queue and idx < len(queue):
            current = queue[idx]
            image = current["image"]
            st.session_state.image_bytes = current["image_bytes"]

            if len(queue) > 1:
                st.info(f"📄 **{idx + 1} / {len(queue)} 枚目**: {current['filename']}")

            # PDFの表裏など複数ページがある場合は並べて表示
            page_images = current["images"]
            if len(page_images) > 1:
                cols = st.columns(len(page_images))
                for i, pg_img in enumerate(page_images):
                    with cols[i]:
                        label = "表面" if i == 0 else "裏面" if i == 1 else f"{i+1}ページ"
                        st.image(pg_img, caption=label, use_container_width=True)
                st.caption(f"📖 {len(page_images)}ページの情報を統合して読み取ります")
            else:
                col_img, col_info = st.columns([1, 1])
                with col_img:
                    st.image(image, caption=current["filename"], use_container_width=True)
                with col_info:
                    st.info(f"**サイズ**: {image.size[0]} × {image.size[1]} px")

            # 現在の画像に対してOCR未実行なら実行（全ページを渡す）
            if not st.session_state.ocr_done:
                _run_ocr(page_images)

# =========================================================================== #
# タブ3: Driveから読取
# =========================================================================== #

with tab_drive:
    st.caption("Google Driveのコース別フォルダから名刺画像を読み取ります。")

    _drive_folder_id = (
        os.getenv("DRIVE_SOURCE_FOLDER_ID")
        or st.secrets.get("DRIVE_SOURCE_FOLDER_ID", "")
        or os.getenv("DRIVE_FOLDER_ID")
        or st.secrets.get("DRIVE_FOLDER_ID", "")
    )

    if not _drive_folder_id:
        st.warning("DRIVE_FOLDER_ID が設定されていません。.env または secrets.toml に設定してください。")
    else:
        if st.button("📂 フォルダ一覧を取得", key="drive_fetch"):
            with st.spinner("Drive フォルダを取得中..."):
                try:
                    subfolders = list_drive_subfolders(_drive_folder_id)
                    st.session_state.drive_subfolders = subfolders
                    if not subfolders:
                        st.info("サブフォルダが見つかりませんでした。")
                except Exception as e:
                    st.error(f"フォルダ取得エラー: {e}")

        if st.session_state.drive_subfolders:
            folder_options = {f["name"]: f for f in st.session_state.drive_subfolders}
            selected_names = st.multiselect(
                "読み取るフォルダを選択してください",
                options=list(folder_options.keys()),
                default=list(folder_options.keys()),
                key="drive_folder_select",
            )

            if st.button("🚀 読取開始", key="drive_start", disabled=not selected_names):
                queue = []
                total_files = 0
                progress_bar = st.progress(0, text="フォルダを読み込み中...")
                for i, name in enumerate(selected_names):
                    folder = folder_options[name]
                    course = _folder_to_course(folder["name"])
                    try:
                        files = list_drive_images(folder["id"])
                    except Exception as e:
                        st.warning(f"⚠️ {folder['name']} の読み込みをスキップ: {e}")
                        continue
                    for f in files:
                        progress_bar.progress(
                            (i + 1) / len(selected_names),
                            text=f"ダウンロード中: {folder['name']} / {f['name']}",
                        )
                        try:
                            raw = download_drive_file(f["id"])
                        except Exception as e:
                            st.warning(f"⚠️ {f['name']} のダウンロードをスキップ: {e}")
                            continue

                        if f["mimeType"] == "application/pdf":
                            imgs = _pdf_to_images(raw)
                        else:
                            imgs = [Image.open(io.BytesIO(raw))]

                        queue.append({
                            "images": imgs,
                            "image": imgs[0],
                            "image_bytes": _to_jpeg_bytes(imgs[0]),
                            "filename": f["name"],
                            "course": course,
                        })
                        total_files += 1

                progress_bar.empty()
                if queue:
                    st.session_state.multi_images = queue
                    st.session_state.multi_idx = 0
                    st.session_state.ocr_done = False
                    st.session_state.card = None
                    st.session_state.drive_course = queue[0].get("course")
                    st.success(f"📥 {total_files} 件のファイルを読み込みました。")
                    st.rerun()
                else:
                    st.info("対象ファイルが見つかりませんでした。")

        # Drive経由のキュー処理（ファイルアップロードと同様にプレビュー・OCR）
        queue = st.session_state.multi_images
        idx = st.session_state.multi_idx
        if queue and idx < len(queue) and queue[idx].get("course") is not None:
            current = queue[idx]
            st.session_state.image_bytes = current["image_bytes"]
            st.session_state.drive_course = current.get("course")

            if len(queue) > 1:
                st.info(f"📄 **{idx + 1} / {len(queue)} 枚目**: {current['filename']}")

            page_images = current["images"]
            if len(page_images) > 1:
                cols = st.columns(len(page_images))
                for i, pg_img in enumerate(page_images):
                    with cols[i]:
                        label = "表面" if i == 0 else "裏面" if i == 1 else f"{i+1}ページ"
                        st.image(pg_img, caption=label, use_container_width=True)
                st.caption(f"📖 {len(page_images)}ページの情報を統合して読み取ります")
            else:
                col_img, col_info = st.columns([1, 1])
                with col_img:
                    st.image(current["image"], caption=current["filename"], use_container_width=True)
                with col_info:
                    st.info(f"**サイズ**: {current['image'].size[0]} × {current['image'].size[1]} px")

            if not st.session_state.ocr_done:
                _run_ocr(page_images)

# --------------------------------------------------------------------------- #
# OCR結果の確認・修正フォーム（カメラ・ファイル両タブ共通）
# --------------------------------------------------------------------------- #

if st.session_state.ocr_done and st.session_state.card is not None and not st.session_state.submitted:
    card = st.session_state.card

    st.markdown("### 内容を確認・修正")
    st.caption("OCR結果を確認し、誤りがあれば修正してから登録してください。")

    with st.form(f"ocr_confirm_form_{fk}"):
        col1, col2 = st.columns(2)

        with col1:
            company_name = st.text_input("企業名 *",         value=card.company_name)
            full_name    = st.text_input("氏名 *",           value=card.full_name)
            title        = st.text_input("役職",             value=card.title)

        with col2:
            email        = st.text_input("メールアドレス *", value=card.email)
            department   = st.text_input("部署",             value=card.department)
            phone        = st.text_input("電話番号",         value=card.phone)

        # Driveフォルダからのコース自動判定
        _auto_course = st.session_state.get("drive_course")
        _course_idx = None
        if _auto_course and _auto_course in COURSE_OPTIONS:
            _course_idx = COURSE_OPTIONS.index(_auto_course)

        course = st.radio(
            "コース",
            COURSE_OPTIONS,
            index=_course_idx,
            key=f"ocr_course_{fk}",
            horizontal=False,
        )

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
                with st.spinner("Google Drive に画像をアップロード中..."):
                    image_url = ""
                    if st.session_state.image_bytes:
                        try:
                            fname = f"{full_name.strip()}_{company_name.strip()}.jpg"
                            image_url = upload_to_drive(st.session_state.image_bytes, fname)
                        except Exception as e:
                            st.warning(f"画像アップロードをスキップしました: {e}")
                with st.spinner("Google Sheetsに書き込み中..."):
                    try:
                        row_num = append_business_card(
                            confirmed_card,
                            source="名刺OCR",
                            image_url=image_url,
                            interest=course or "",
                        )
                        # 複数ファイルキューに次があれば進む
                        queue = st.session_state.get("multi_images", [])
                        idx = st.session_state.get("multi_idx", 0)
                        if queue and idx + 1 < len(queue):
                            st.session_state.multi_idx += 1
                            st.session_state.ocr_done = False
                            st.session_state.card = None
                            st.session_state.image_bytes = None
                            # 次のアイテムのコース自動設定（Drive経由の場合）
                            next_item = queue[idx + 1]
                            st.session_state.drive_course = next_item.get("course")
                            st.toast(f"✅ {idx + 1}/{len(queue)} 枚目を登録しました。次へ進みます...")
                            st.rerun()
                        else:
                            st.session_state.submitted = True
                            st.rerun()
                    except FileNotFoundError as e:
                        st.error(f"認証エラー: {e}")
                    except ValueError as e:
                        st.error(f"設定エラー: {e}")
                    except Exception as e:
                        st.error(f"書き込みエラー: {e}")

# =========================================================================== #
# タブ3: 手動入力
# =========================================================================== #

with tab_manual:

    st.markdown("### 情報を手動で入力")
    st.caption("名刺がない場合や、直接入力したい場合にご利用ください。")

    with st.form(f"manual_form_{fk}"):
        col1, col2 = st.columns(2)

        with col1:
            m_company_name = st.text_input("企業名 *",         key=f"m_company_{fk}")
            m_full_name    = st.text_input("氏名 *",           key=f"m_fullname_{fk}")
            m_title        = st.text_input("役職",             key=f"m_title_{fk}")

        with col2:
            m_email        = st.text_input("メールアドレス *", key=f"m_email_{fk}")
            m_department   = st.text_input("部署",             key=f"m_dept_{fk}")
            m_phone        = st.text_input("電話番号",         key=f"m_phone_{fk}")

        m_course = st.radio(
            "コース",
            COURSE_OPTIONS,
            index=None,
            key=f"m_course_{fk}",
            horizontal=False,
        )

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
                        row_num = append_business_card(
                            manual_card, source="手動入力", interest=m_course or ""
                        )
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
    st.session_state.image_bytes = None
    st.session_state.multi_images = []
    st.session_state.multi_idx = 0
    st.session_state.multi_upload_key = None
    st.session_state.drive_subfolders = []
    st.session_state.drive_course = None
    st.session_state.dup_cache = {}
    st.session_state.form_key += 1  # フォームキーを更新してフィールドをリセット
    st.rerun()
