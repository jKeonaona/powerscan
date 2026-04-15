"""Fire-and-forget email notifications for PowerScan feedback submissions.

Uses the Resend HTTP API (https://api.resend.com/emails). Config comes
from app.config, which is loaded from .env via config.py:

    RESEND_API_KEY  — Resend API key (Bearer token)
    ADMIN_EMAIL     — recipient for feedback notifications
    APP_PUBLIC_URL  — shown at the bottom of the email body

If either RESEND_API_KEY or ADMIN_EMAIL is missing, we log a single-line
warning and return — the feedback row is still saved.
"""
import threading
import traceback

import requests

from models import db, Feedback

RESEND_ENDPOINT = "https://api.resend.com/emails"
FROM_ADDRESS = "PowerScan <powerscan@notify.ccctrainingonline.com>"
HTTP_TIMEOUT = 15  # seconds


def _send_sync(app, feedback_id):
    with app.app_context():
        try:
            fb = db.session.get(Feedback, feedback_id)
            if not fb:
                return

            api_key = (app.config.get("RESEND_API_KEY") or "").strip()
            admin_email = (app.config.get("ADMIN_EMAIL") or "").strip()
            app_url = (app.config.get("APP_PUBLIC_URL") or "https://powerscan.ccctrainingonline.com").strip()

            if not api_key or not admin_email:
                missing = []
                if not api_key:
                    missing.append("RESEND_API_KEY")
                if not admin_email:
                    missing.append("ADMIN_EMAIL")
                print(
                    f"[powerscan] feedback email skipped for #{fb.id}: missing config {missing}",
                    flush=True,
                )
                return

            user = fb.user
            user_label = user.username if user else "(unknown user)"
            user_email = user.email if user else ""

            snippet = " ".join((fb.description or "").split())[:60]
            subject = f"PowerScan Feedback — {fb.type}: {snippet}"

            body_lines = [
                f"Type: {fb.type}",
                f"From: {user_label} <{user_email or 'unknown'}>",
                f"Page: {fb.page or '(not provided)'}",
                f"Submitted: {fb.created_at.strftime('%Y-%m-%d %H:%M UTC')}",
                "",
                "Description:",
                fb.description or "(empty)",
                "",
                f"Log in to review: {app_url}",
            ]
            body = "\n".join(body_lines)

            payload = {
                "from": FROM_ADDRESS,
                "to": [admin_email],
                "subject": subject,
                "text": body,
            }
            if user_email and "@" in str(user_email):
                payload["reply_to"] = user_email

            headers = {
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
            }

            resp = requests.post(
                RESEND_ENDPOINT,
                json=payload,
                headers=headers,
                timeout=HTTP_TIMEOUT,
            )

            if 200 <= resp.status_code < 300:
                try:
                    resend_id = resp.json().get("id", "?")
                except Exception:
                    resend_id = "?"
                print(
                    f"[powerscan] feedback email sent via Resend to {admin_email} "
                    f"for feedback #{fb.id} (resend id {resend_id})",
                    flush=True,
                )
            else:
                body_preview = (resp.text or "")[:300]
                print(
                    f"[powerscan] feedback email failed for #{fb.id}: "
                    f"Resend returned HTTP {resp.status_code} — {body_preview}",
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


def _send_reply_sync(app, feedback_id):
    with app.app_context():
        try:
            fb = db.session.get(Feedback, feedback_id)
            if not fb:
                return

            api_key = (app.config.get("RESEND_API_KEY") or "").strip()
            app_url = (app.config.get("APP_PUBLIC_URL") or "https://powerscan.ccctrainingonline.com").strip()

            user = fb.user
            if not user or not user.email or "@" not in str(user.email):
                print(f"[powerscan] reply email skipped for #{fb.id}: no user email", flush=True)
                return

            if not api_key:
                print(f"[powerscan] reply email skipped for #{fb.id}: missing RESEND_API_KEY", flush=True)
                return

            subject = "PowerScan \u2014 Response to your feedback"

            sep = "\u2500" * 60
            body_lines = [
                f"Hi {user.username},",
                "",
                "Thank you for your feedback. Here is a response from the PowerScan team:",
                "",
                sep,
                "Your original feedback:",
                f"  Type: {fb.type}",
                f"  Submitted: {fb.created_at.strftime('%Y-%m-%d %H:%M UTC')}",
                "",
                fb.description or "(no description)",
                sep,
                "",
                "Admin reply:",
                fb.admin_reply or "(no reply text)",
                "",
                f"View on the feedback dashboard: {app_url}/admin/feedback",
            ]
            body = "\n".join(body_lines)

            payload = {
                "from": FROM_ADDRESS,
                "to": [user.email],
                "subject": subject,
                "text": body,
            }

            headers = {
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
            }

            resp = requests.post(
                RESEND_ENDPOINT,
                json=payload,
                headers=headers,
                timeout=HTTP_TIMEOUT,
            )

            if 200 <= resp.status_code < 300:
                try:
                    resend_id = resp.json().get("id", "?")
                except Exception:
                    resend_id = "?"
                print(
                    f"[powerscan] reply email sent via Resend to {user.email} "
                    f"for feedback #{fb.id} (resend id {resend_id})",
                    flush=True,
                )
            else:
                body_preview = (resp.text or "")[:300]
                print(
                    f"[powerscan] reply email failed for #{fb.id}: "
                    f"Resend returned HTTP {resp.status_code} — {body_preview}",
                    flush=True,
                )

        except Exception as e:
            print(f"[powerscan] reply email failed: {e}", flush=True)
            traceback.print_exc()


def send_reply_email_async(app, feedback_id):
    """Send an admin-reply notification to the feedback submitter; never raises."""
    try:
        thread = threading.Thread(
            target=_send_reply_sync, args=(app, feedback_id), daemon=True,
        )
        thread.start()
    except Exception as e:
        print(f"[powerscan] failed to spawn reply email thread: {e}", flush=True)
