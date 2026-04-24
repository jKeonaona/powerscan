"""Fire-and-forget email notifications for PowerScan.

Uses the Resend HTTP API (https://api.resend.com/emails). Config comes
from app.config, which is loaded from .env via config.py:

    RESEND_API_KEY  — Resend API key (Bearer token)
    APP_PUBLIC_URL  — base URL included in outbound emails
"""
import threading
import traceback

import requests

from models import db, PasswordResetToken

RESEND_ENDPOINT = "https://api.resend.com/emails"
FROM_ADDRESS = "PowerScan <powerscan@notify.ccctrainingonline.com>"
HTTP_TIMEOUT = 15  # seconds


def _send_password_reset_sync(app, token_id):
    with app.app_context():
        try:
            tok = db.session.get(PasswordResetToken, token_id)
            if not tok:
                return

            api_key = (app.config.get("RESEND_API_KEY") or "").strip()
            app_url = (app.config.get("APP_PUBLIC_URL") or "https://powerscan.ccctrainingonline.com").strip()

            user = tok.user
            if not user or not user.email or "@" not in str(user.email):
                print(f"[powerscan] reset email skipped for token #{tok.id}: no user email", flush=True)
                return

            if not api_key:
                print(f"[powerscan] reset email skipped for token #{tok.id}: missing RESEND_API_KEY", flush=True)
                return

            reset_url = f"{app_url}/reset-password/{tok.token}"
            subject = "PowerScan — Password Reset Request"
            body_lines = [
                f"Hi {user.username},",
                "",
                "We received a request to reset your PowerScan password.",
                "Click the link below to choose a new password:",
                "",
                reset_url,
                "",
                "This link expires in 1 hour.",
                "If you did not request a password reset, you can safely ignore this email.",
            ]
            body = "\n".join(body_lines)

            payload = {"from": FROM_ADDRESS, "to": [user.email], "subject": subject, "text": body}
            headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}

            resp = requests.post(RESEND_ENDPOINT, json=payload, headers=headers, timeout=HTTP_TIMEOUT)
            if 200 <= resp.status_code < 300:
                try:
                    resend_id = resp.json().get("id", "?")
                except Exception:
                    resend_id = "?"
                print(
                    f"[powerscan] password reset email sent to {user.email} "
                    f"for token #{tok.id} (resend id {resend_id})",
                    flush=True,
                )
            else:
                print(
                    f"[powerscan] password reset email failed for token #{tok.id}: "
                    f"HTTP {resp.status_code} — {(resp.text or '')[:200]}",
                    flush=True,
                )
        except Exception as e:
            print(f"[powerscan] password reset email failed: {e}", flush=True)
            traceback.print_exc()


def send_password_reset_email_async(app, token_id):
    """Send a password-reset link to the user; never raises."""
    try:
        thread = threading.Thread(
            target=_send_password_reset_sync, args=(app, token_id), daemon=True,
        )
        thread.start()
    except Exception as e:
        print(f"[powerscan] failed to spawn password reset email thread: {e}", flush=True)
