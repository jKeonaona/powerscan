import os

from dotenv import load_dotenv

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
_ENV_PATH = os.path.join(BASE_DIR, ".env")
load_dotenv(_ENV_PATH)

_key = os.environ.get("ANTHROPIC_API_KEY", "")
_env_loaded = os.path.isfile(_ENV_PATH)
print(
    f"[powerscan] .env {'loaded from ' + _ENV_PATH if _env_loaded else 'NOT FOUND at ' + _ENV_PATH}",
    flush=True,
)
print(f"[powerscan] ANTHROPIC_API_KEY loaded: {_key[:10] if _key else '(empty)'}", flush=True)


class Config:
    # Fixed fallback (NOT random) so sessions survive restarts even without .env.
    SECRET_KEY = os.environ.get("SECRET_KEY", "powerscan-dev-key-change-in-production")
    SQLALCHEMY_DATABASE_URI = f"sqlite:///{os.path.join(BASE_DIR, 'instance', 'powerscan.db')}"
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
    PROCESSED_FOLDER = os.path.join(BASE_DIR, "processed")
    REPORTS_FOLDER = os.path.join(BASE_DIR, "reports_output")
    MAX_CONTENT_LENGTH = 100 * 1024 * 1024
    ANTHROPIC_API_KEY = _key

    # ── Email notifications (Resend HTTP API) ───────────────────
    RESEND_API_KEY = os.environ.get("RESEND_API_KEY", "")
    ADMIN_EMAIL = os.environ.get("ADMIN_EMAIL", "")
    APP_PUBLIC_URL = os.environ.get("APP_PUBLIC_URL", "https://powerscan.ccctrainingonline.com")
