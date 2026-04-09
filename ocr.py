import os
import threading
import time
from datetime import datetime, timezone

import cv2
import numpy as np
import pytesseract
from pdf2image import convert_from_path

from models import db, Drawing, DrawingPage

POLL_INTERVAL = 2  # seconds between queue checks


def configure_tesseract(tesseract_cmd):
    pytesseract.pytesseract.tesseract_cmd = tesseract_cmd


def preprocess_image(image_path):
    """Apply OpenCV preprocessing to improve OCR accuracy on engineering drawings."""
    img = cv2.imread(image_path)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    thresh = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 31, 10
    )
    denoised = cv2.fastNlMeansDenoising(thresh, h=10)
    coords = np.column_stack(np.where(denoised > 0))
    if len(coords) > 5:
        angle = cv2.minAreaRect(coords)[-1]
        if angle < -45:
            angle = -(90 + angle)
        else:
            angle = -angle
        if abs(angle) > 0.5:
            h, w = denoised.shape
            center = (w // 2, h // 2)
            matrix = cv2.getRotationMatrix2D(center, angle, 1.0)
            denoised = cv2.warpAffine(
                denoised, matrix, (w, h),
                flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE
            )
    return denoised


def _process_one(app, drawing_id):
    """Convert a single drawing's PDF pages to JPEGs, run OCR, store text."""
    with app.app_context():
        drawing = db.session.get(Drawing, drawing_id)
        if not drawing:
            return

        drawing.status = "processing"
        db.session.commit()

        pdf_path = os.path.join(app.config["UPLOAD_FOLDER"], drawing.filename)
        drawing_dir = os.path.join(app.config["PROCESSED_FOLDER"], str(drawing.id))
        os.makedirs(drawing_dir, exist_ok=True)

        try:
            images = convert_from_path(pdf_path, dpi=300)

            drawing.total_pages = len(images)
            drawing.pages_processed = 0
            db.session.commit()

            for page_num, image in enumerate(images, start=1):
                image_filename = f"page_{page_num}.jpg"
                image_path = os.path.join(drawing_dir, image_filename)
                image.save(image_path, "JPEG", quality=95)

                processed = preprocess_image(image_path)

                custom_config = r"--oem 3 --psm 6"
                text = pytesseract.image_to_string(processed, config=custom_config)

                page = DrawingPage(
                    drawing_id=drawing.id,
                    page_number=page_num,
                    image_path=f"{drawing.id}/{image_filename}",
                    extracted_text=text.strip(),
                    processed_at=datetime.now(timezone.utc),
                )
                db.session.add(page)

                drawing.pages_processed = page_num
                db.session.commit()

            drawing.status = "completed"
            db.session.commit()
            print(f"OCR completed for drawing {drawing_id} ({drawing.original_filename})")

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
