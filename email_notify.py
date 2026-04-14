"""Fire-and-forget email notifications for PowerScan feedback submissions.

SMTP settings come from app.config (loaded from .env via config.py):
    SMTP_SERVER, SMTP_PORT, SMTP_USER, SMTP_PASSWORD, ADMIN_EMAIL
    APP_PUBLIC_URL (used in the 'Log in to review' line)

If any required field is missing, we log a single-line warning and
return cleanly — the feedback row is still saved either way.
"""
import smtplib
import ssl
import threading
import traceback
from email.message import EmailMessage

from models import db, Feedback


def _send_sync(app, feedback_id):
    with app.app_context():
        try:
            fb = db.session.get(Feedback, feedback_id)
            if not fb:
                return

            admin_email = (app.config.get("ADMIN_EMAIL") or "").strip()
            smtp_server = (app.config.get("SMTP_SERVER") or "").strip()
            smtp_port = int(app.config.get("SMTP_PORT") or 0)
            smtp_user = (app.config.get("SMTP_USER") or "").strip()
            smtp_password = app.config.get("SMTP_PASSWORD") or ""
            app_url = (app.config.get("APP_PUBLIC_URL") or "https://powerscan.ccctrainingonline.com").strip()

            missing = [
                name for name, val in (
                    ("ADMIN_EMAIL", admin_email),
                    ("SMTP_SERVER", smtp_server),
                    ("SMTP_PORT", smtp_port),
                    ("SMTP_USER", smtp_user),
                    ("SMTP_PASSWORD", smtp_password),
                ) if not val
            ]
            if missing:
                print(
                    f"[powerscan] feedback email skipped for #{fb.id}: "
                    f"missing config {missing}",
                    flush=True,
                )
                return

            user = fb.user
            user_label = user.username if user else "(unknown user)"
            user_email = user.email if user else "(unknown)"

            snippet = " ".join((fb.description or "").split())[:60]
            subject = f"PowerScan Feedback — {fb.type}: {snippet}"

            body_lines = [
                f"Type: {fb.type}",
                f"From: {user_label} <{user_email}>",
                f"Page: {fb.page or '(not provided)'}",
                f"Submitted: {fb.created_at.strftime('%Y-%m-%d %H:%M UTC')}",
                "",
                "Description:",
                fb.description or "(empty)",
                "",
                f"Log in to review: {app_url}",
            ]
            body = "\n".join(body_lines)

            msg = EmailMessage()
            msg["Subject"] = subject
            msg["From"] = smtp_user
            msg["To"] = admin_email
            if user_email and "@" in str(user_email):
                msg["Reply-To"] = user_email
            msg.set_content(body)

            if smtp_port == 465:
                context = ssl.create_default_context()
                with smtplib.SMTP_SSL(smtp_server, smtp_port, context=context, timeout=30) as server:
                    server.login(smtp_user, smtp_password)
                    server.send_message(msg)
            else:
                with smtplib.SMTP(smtp_server, smtp_port, timeout=30) as server:
                    server.ehlo()
                    server.starttls(context=ssl.create_default_context())
                    server.ehlo()
                    server.login(smtp_user, smtp_password)
                    server.send_message(msg)

            print(
                f"[powerscan] feedback email sent to {admin_email} for feedback #{fb.id}",
                flush=True,
            )
        except Exception as e:
            print(f"[powerscan] feedback email failed: {e}", flush=True)
            traceback.print_exc()


def send_feedback_email_async(app, feedback_id):
    """Kick off a daemon thread to send the notification; never raises."""
    try:
        thread = threading.Thread(
            target=_send_sync, args=(app, feedback_id), daemon=True,
        )
        thread.start()
    except Exception as e:
        print(f"[powerscan] failed to spawn feedback email thread: {e}", flush=True)
