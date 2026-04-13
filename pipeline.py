import os
import threading
import time
from datetime import datetime, timezone

from pdf2image import convert_from_path, pdfinfo_from_path

from models import db, Drawing, DrawingPage

POLL_INTERVAL = 2
PAGE_DPI = 200


def _convert_one(app, drawing_id):
    with app.app_context():
        drawing = db.session.get(Drawing, drawing_id)
        if not drawing:
            return

        drawing.status = "processing"
        drawing.pages_processed = 0
        DrawingPage.query.filter_by(drawing_id=drawing.id).delete()
        db.session.commit()

        pdf_path = os.path.join(app.config["UPLOAD_FOLDER"], drawing.filename)
        drawing_dir = os.path.join(app.config["PROCESSED_FOLDER"], str(drawing.id))
        os.makedirs(drawing_dir, exist_ok=True)

        try:
            info = pdfinfo_from_path(pdf_path)
            total = info.get("Pages", 0)
            drawing.total_pages = total
            db.session.commit()

            print(f"Drawing {drawing_id}: converting {total} pages @ {PAGE_DPI} DPI")

            for page_num in range(1, total + 1):
                images = convert_from_path(
                    pdf_path, dpi=PAGE_DPI,
                    first_page=page_num, last_page=page_num,
                )
                image = images[0]
                image_filename = f"page_{page_num}.jpg"
                image_path = os.path.join(drawing_dir, image_filename)
                image.save(image_path, "JPEG", quality=90)
                del images, image

                page = DrawingPage(
                    drawing_id=drawing.id,
                    page_number=page_num,
                    image_path=f"{drawing.id}/{image_filename}",
                    processed_at=datetime.now(timezone.utc),
                )
                db.session.add(page)
                drawing.pages_processed = page_num
                db.session.commit()

            drawing.status = "ready"
            db.session.commit()
            print(f"Drawing {drawing_id} ready ({drawing.original_filename})")

        except Exception as e:
            drawing.status = "failed"
            db.session.commit()
            print(f"Conversion failed for drawing {drawing_id}: {e}")


def _worker_loop(app):
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
                    _convert_one(app, drawing.id)
                else:
                    time.sleep(POLL_INTERVAL)
        except Exception as e:
            print(f"Pipeline worker error: {e}")
            time.sleep(POLL_INTERVAL)


def start_worker(app):
    thread = threading.Thread(target=_worker_loop, args=(app,), daemon=True)
    thread.start()
    print("Pipeline worker started")
