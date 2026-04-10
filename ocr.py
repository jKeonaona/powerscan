import base64
import os
import threading
import time
from datetime import datetime, timezone

import anthropic
import cv2
import numpy as np
import pytesseract
from pdf2image import convert_from_path, pdfinfo_from_path

from models import db, Drawing, DrawingPage

POLL_INTERVAL = 2  # seconds between queue checks
COOLDOWN_BETWEEN_DRAWINGS = 30  # seconds between finishing one drawing and starting the next

# Auto-escalating DPI passes (low → medium → high)
DPI_PASSES = [100, 150, 300]


def configure_tesseract(tesseract_cmd):
    pytesseract.pytesseract.tesseract_cmd = tesseract_cmd


def preprocess_image(image_path):
    """Apply OpenCV preprocessing optimized for old engineering drawings."""
    img = cv2.imread(image_path)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # 1) CLAHE — boost contrast on faded/uneven scans
    clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8, 8))
    enhanced = clahe.apply(gray)

    # 2) Bilateral filter — remove noise while preserving text edges
    filtered = cv2.bilateralFilter(enhanced, 9, 75, 75)

    # 3) Otsu binarization — automatically picks the best threshold
    _, binary = cv2.threshold(filtered, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

    # 4) Morphological cleanup — close small gaps in text, remove speckle
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2))
    cleaned = cv2.morphologyEx(binary, cv2.MORPH_CLOSE, kernel)
    cleaned = cv2.morphologyEx(cleaned, cv2.MORPH_OPEN, kernel)

    # 5) Aggressive denoise on the binary image
    denoised = cv2.fastNlMeansDenoising(cleaned, h=15)

    # 6) Deskew
    coords = np.column_stack(np.where(denoised > 0))
    if len(coords) > 100:
        angle = cv2.minAreaRect(coords)[-1]
        if angle < -45:
            angle = -(90 + angle)
        else:
            angle = -angle
        if abs(angle) > 0.3:
            h, w = denoised.shape
            center = (w // 2, h // 2)
            matrix = cv2.getRotationMatrix2D(center, angle, 1.0)
            denoised = cv2.warpAffine(
                denoised, matrix, (w, h),
                flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE
            )

    return denoised


def _score_text(text):
    """Score OCR text quality — higher is better."""
    if not text:
        return 0
    words = text.split()
    if not words:
        return 0
    # Favor longer text with more real words (3+ chars, mostly alpha)
    real_words = sum(1 for w in words if len(w) >= 3 and sum(c.isalpha() for c in w) / len(w) > 0.5)
    # Penalize high ratio of garbage characters
    alpha_ratio = sum(c.isalnum() or c.isspace() for c in text) / len(text)
    return real_words * alpha_ratio


def _extract_with_claude_vision(image_path, api_key):
    """Send page image to Claude Vision API to extract text."""
    if not api_key:
        return ""

    try:
        with open(image_path, "rb") as f:
            image_data = base64.standard_b64encode(f.read()).decode("utf-8")

        client = anthropic.Anthropic(api_key=api_key)
        message = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=4096,
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": "image/jpeg",
                                "data": image_data,
                            },
                        },
                        {
                            "type": "text",
                            "text": "Extract ALL text visible in this engineering drawing. "
                                    "Include title blocks, notes, dimensions, labels, part numbers, "
                                    "revision info, and any other text. Preserve the layout as much "
                                    "as possible. Output only the extracted text, nothing else.",
                        },
                    ],
                }
            ],
        )
        return message.content[0].text.strip()
    except Exception as e:
        print(f"Claude Vision extraction failed: {e}")
        return ""


def _run_pass(app, drawing, pdf_path, drawing_dir, dpi, use_claude, api_key):
    """Run one OCR pass at a given DPI, replacing all existing pages."""
    DrawingPage.query.filter_by(drawing_id=drawing.id).delete()
    drawing.ocr_dpi = dpi
    drawing.pages_processed = 0
    db.session.commit()

    info = pdfinfo_from_path(pdf_path)
    total = info.get("Pages", 0)
    drawing.total_pages = total
    db.session.commit()

    print(f"  Pass @ {dpi} DPI — {total} pages{' + Claude Vision' if use_claude else ''}")

    for page_num in range(1, total + 1):
        images = convert_from_path(
            pdf_path, dpi=dpi,
            first_page=page_num, last_page=page_num,
        )
        image = images[0]

        image_filename = f"page_{page_num}.jpg"
        image_path = os.path.join(drawing_dir, image_filename)
        image.save(image_path, "JPEG", quality=95)
        del images, image

        # --- Tesseract OCR ---
        processed = preprocess_image(image_path)
        custom_config = r"--oem 3 --psm 6"
        tesseract_text = pytesseract.image_to_string(processed, config=custom_config).strip()
        del processed

        best_text = tesseract_text
        source = "tesseract"

        # --- Claude Vision OCR (final pass only) ---
        if use_claude:
            claude_text = _extract_with_claude_vision(image_path, api_key)
            tesseract_score = _score_text(tesseract_text)
            claude_score = _score_text(claude_text)
            if claude_text and claude_score > tesseract_score:
                best_text = claude_text
                source = "claude"
            print(f"    Page {page_num}: tesseract={tesseract_score:.0f}, claude={claude_score:.0f} -> {source}")

        page = DrawingPage(
            drawing_id=drawing.id,
            page_number=page_num,
            image_path=f"{drawing.id}/{image_filename}",
            extracted_text=best_text,
            processed_at=datetime.now(timezone.utc),
        )
        db.session.add(page)

        drawing.pages_processed = page_num
        db.session.commit()


def _process_one(app, drawing_id):
    """Auto-escalate through DPI passes: 100 -> 150 -> 300 (with Claude Vision)."""
    with app.app_context():
        drawing = db.session.get(Drawing, drawing_id)
        if not drawing:
            return

        drawing.status = "processing"
        db.session.commit()

        pdf_path = os.path.join(app.config["UPLOAD_FOLDER"], drawing.filename)
        drawing_dir = os.path.join(app.config["PROCESSED_FOLDER"], str(drawing.id))
        os.makedirs(drawing_dir, exist_ok=True)

        api_key = app.config.get("ANTHROPIC_API_KEY", "")

        try:
            # Manual reprocess at a specific DPI — single pass, no escalation
            if drawing.ocr_dpi and drawing.ocr_dpi > 0:
                dpi = drawing.ocr_dpi
                use_claude = (dpi >= 300)
                print(f"Drawing {drawing_id}: reprocess @ {dpi} DPI")
                _run_pass(app, drawing, pdf_path, drawing_dir, dpi, use_claude, api_key)
            else:
                # Auto-escalating passes: 100 -> 150 -> 300
                print(f"Drawing {drawing_id}: auto-escalating {DPI_PASSES}")
                for dpi in DPI_PASSES:
                    use_claude = (dpi == DPI_PASSES[-1])
                    _run_pass(app, drawing, pdf_path, drawing_dir, dpi, use_claude, api_key)

            drawing.status = "completed"
            db.session.commit()
            print(f"OCR completed for drawing {drawing_id} ({drawing.original_filename}) @ {drawing.ocr_dpi} DPI")

        except Exception as e:
            drawing.status = "failed"
            db.session.commit()
            print(f"OCR processing failed for drawing {drawing_id}: {e}")


def _worker_loop(app):
    """Continuously poll for pending drawings and process them one at a time."""
    # On startup, recover any drawings stuck in 'processing' from a previous crash
    with app.app_context():
        stuck = Drawing.query.filter_by(status="processing").all()
        for d in stuck:
            print(f"Recovering stuck drawing {d.id} ({d.original_filename})")
            d.status = "pending"
        if stuck:
            db.session.commit()

    while True:
        try:
            with app.app_context():
                drawing = (
                    Drawing.query
                    .filter_by(status="pending")
                    .order_by(Drawing.created_at)
                    .first()
                )
                if drawing:
                    _process_one(app, drawing.id)
                    time.sleep(COOLDOWN_BETWEEN_DRAWINGS)
                else:
                    time.sleep(POLL_INTERVAL)
        except Exception as e:
            print(f"OCR worker error: {e}")
            time.sleep(POLL_INTERVAL)


def start_worker(app):
    """Launch the background OCR worker thread."""
    thread = threading.Thread(target=_worker_loop, args=(app,), daemon=True)
    thread.start()
    print("OCR background worker started")
