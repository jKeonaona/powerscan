import os

from dotenv import load_dotenv

load_dotenv()

BASE_DIR = os.path.abspath(os.path.dirname(__file__))

_key = os.environ.get("ANTHROPIC_API_KEY", "")
print(f"[powerscan] ANTHROPIC_API_KEY loaded: {_key[:10] if _key else '(empty)'}", flush=True)


class Config:
    SECRET_KEY = os.environ.get("SECRET_KEY", "powerscan-dev-key-change-in-production")
    SQLALCHEMY_DATABASE_URI = f"sqlite:///{os.path.join(BASE_DIR, 'instance', 'powerscan.db')}"
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
    PROCESSED_FOLDER = os.path.join(BASE_DIR, "processed")
    REPORTS_FOLDER = os.path.join(BASE_DIR, "reports_output")
    MAX_CONTENT_LENGTH = 100 * 1024 * 1024
    ANTHROPIC_API_KEY = _key
