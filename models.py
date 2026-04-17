from datetime import datetime, timezone
from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash

db = SQLAlchemy()

ROLE_SUPERADMIN = "superadmin"
ROLE_ADMIN = "admin"
ROLE_USER = "user"
ROLES = [ROLE_SUPERADMIN, ROLE_ADMIN, ROLE_USER]

DOC_TYPES = ["Drawing", "Contract", "Specification", "Bid Doc", "Addendum", "Estimation Notes", "Other"]
DEFAULT_DOC_TYPE = "Drawing"

FEEDBACK_TYPES = ["Idea", "Bug Report", "Question", "Feature Request"]
FEEDBACK_STATUSES = ["Pending", "Reviewed", "Adopted", "Declined"]
DEFAULT_FEEDBACK_STATUS = "Pending"


class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False)
    password_hash = db.Column(db.String(256), nullable=False)
    role = db.Column(db.String(20), nullable=False, default=ROLE_USER)
    company_id = db.Column(db.Integer, db.ForeignKey("company.id"), nullable=True)
    must_change_password = db.Column(db.Boolean, nullable=False, default=False)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    company = db.relationship("Company", backref="users")

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

    @property
    def is_superadmin(self):
        return self.role == ROLE_SUPERADMIN

    @property
    def is_admin(self):
        return self.role in (ROLE_SUPERADMIN, ROLE_ADMIN)


class Company(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200), unique=True, nullable=False)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    projects = db.relationship("Project", backref="company", cascade="all, delete-orphan")


class Project(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200), nullable=False)
    description = db.Column(db.Text)
    company_id = db.Column(db.Integer, db.ForeignKey("company.id"), nullable=False)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    work_scope = db.Column(db.Text, nullable=True)   # JSON array of selected scope items
    scope_details = db.Column(db.Text, nullable=True)

    drawings = db.relationship("Drawing", backref="project", cascade="all, delete-orphan")

    @property
    def work_scope_list(self):
        if not self.work_scope:
            return []
        try:
            import json
            return json.loads(self.work_scope)
        except Exception:
            return []


class Drawing(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    filename = db.Column(db.String(300), nullable=False)
    original_filename = db.Column(db.String(300), nullable=False)
    project_id = db.Column(db.Integer, db.ForeignKey("project.id"), nullable=False)
    uploaded_by = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    doc_type = db.Column(db.String(40), default=DEFAULT_DOC_TYPE, nullable=False)
    total_pages = db.Column(db.Integer, default=0)
    pages_processed = db.Column(db.Integer, default=0)
    status = db.Column(db.String(20), default="pending")  # pending, processing, ready, failed
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    uploader = db.relationship("User", backref="drawings")
    pages = db.relationship("DrawingPage", backref="drawing", cascade="all, delete-orphan")


class DrawingPage(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    drawing_id = db.Column(db.Integer, db.ForeignKey("drawing.id"), nullable=False)
    page_number = db.Column(db.Integer, nullable=False)
    image_path = db.Column(db.String(500), nullable=False)
    processed_at = db.Column(db.DateTime, nullable=True)


class SearchHistory(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey("project.id"), nullable=False)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    query = db.Column(db.Text, nullable=False)
    answer = db.Column(db.Text, nullable=False, default="")
    doc_type_filter = db.Column(db.String(40), nullable=True)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    project = db.relationship("Project", backref=db.backref("search_history", cascade="all, delete-orphan"))
    user = db.relationship("User")


class Feedback(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    type = db.Column(db.String(40), nullable=False)
    page = db.Column(db.String(500), nullable=True)
    description = db.Column(db.Text, nullable=False)
    status = db.Column(db.String(20), nullable=False, default=DEFAULT_FEEDBACK_STATUS)
    admin_notes = db.Column(db.Text, nullable=True)
    admin_reply = db.Column(db.Text, nullable=True)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    user = db.relationship("User")


class LaborRate(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    category = db.Column(db.String(100), nullable=False)
    craft_type = db.Column(db.String(100), nullable=False)
    region = db.Column(db.String(100), nullable=False, default="General")
    hourly_cost = db.Column(db.Float, nullable=False)
    effective_date = db.Column(db.Date, nullable=True)
    expiry_date = db.Column(db.Date, nullable=True)
    version = db.Column(db.Integer, nullable=False, default=1)
    uploaded_by = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    active = db.Column(db.Boolean, nullable=False, default=True)

    uploader = db.relationship("User", foreign_keys=[uploaded_by])


class InsuranceRate(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    category = db.Column(db.String(100), nullable=False)
    rate_type = db.Column(db.String(100), nullable=False)
    rate_percent = db.Column(db.Float, nullable=False)
    effective_date = db.Column(db.Date, nullable=True)
    expiry_date = db.Column(db.Date, nullable=True)
    version = db.Column(db.Integer, nullable=False, default=1)
    notes = db.Column(db.Text, nullable=True)
    uploaded_by = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    active = db.Column(db.Boolean, nullable=False, default=True)

    uploader = db.relationship("User", foreign_keys=[uploaded_by])


class PasswordResetToken(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    token = db.Column(db.String(100), unique=True, nullable=False, index=True)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    used = db.Column(db.Boolean, nullable=False, default=False)

    user = db.relationship("User")


class LoginEvent(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    ip_address = db.Column(db.String(45), nullable=True)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    user = db.relationship("User")


class Report(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey("project.id"), nullable=False)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    template_id = db.Column(db.String(40), nullable=False)
    template_name = db.Column(db.String(100), nullable=False)
    custom_prompt = db.Column(db.Text, nullable=True)
    status = db.Column(db.String(20), nullable=False, default="generating")  # generating, ready, failed
    file_path = db.Column(db.String(300), nullable=True)
    error_message = db.Column(db.Text, nullable=True)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    completed_at = db.Column(db.DateTime, nullable=True)

    project = db.relationship("Project", backref=db.backref("reports", cascade="all, delete-orphan"))
    user = db.relationship("User")
