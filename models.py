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


PROJECT_STATUSES = ["Active", "On Hold", "Complete", "Archived"]

TAKEOFF_STATUSES = ["Draft", "Final"]


class Project(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200), nullable=False)
    description = db.Column(db.Text)
    company_id = db.Column(db.Integer, db.ForeignKey("company.id"), nullable=False)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    work_scope = db.Column(db.Text, nullable=True)   # JSON array of selected scope items
    scope_details = db.Column(db.Text, nullable=True)
    status = db.Column(db.String(32), nullable=False, default="Active")
    archived_at = db.Column(db.DateTime, nullable=True)
    bid_date = db.Column(db.Date, nullable=True)

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


class Takeoff(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey("project.id"), nullable=False)
    name = db.Column(db.String(200), nullable=False, default="New Takeoff")
    status = db.Column(db.String(32), nullable=False, default="Draft")
    revision_note = db.Column(db.Text, nullable=True)
    submitted_amount = db.Column(db.Numeric(14, 2), nullable=True)
    scopes = db.Column(db.Text, nullable=True)  # JSON list of scope strings
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc), nullable=False)
    created_by_user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=True)

    project = db.relationship("Project", backref=db.backref("takeoffs", cascade="all, delete-orphan"))
    created_by = db.relationship("User", foreign_keys=[created_by_user_id])

    # ── Project Parameters ─────────────────────────────────────────────────────
    deck_area_sf = db.Column(db.Numeric(12, 2), nullable=True)
    blast_level = db.Column(db.String(32), nullable=True)
    abrasive_type = db.Column(db.String(64), nullable=True)
    abrasive_lb_per_sf = db.Column(db.Numeric(10, 3), nullable=True)

    # ── Materials (6 coating layers × 3 fields) ────────────────────────────────
    primer_vol_pct = db.Column(db.Numeric(6, 2), nullable=True)
    primer_mils = db.Column(db.Numeric(6, 2), nullable=True)
    primer_gal = db.Column(db.Numeric(10, 2), nullable=True)
    second_primer_vol_pct = db.Column(db.Numeric(6, 2), nullable=True)
    second_primer_mils = db.Column(db.Numeric(6, 2), nullable=True)
    second_primer_gal = db.Column(db.Numeric(10, 2), nullable=True)
    stripe_prime_vol_pct = db.Column(db.Numeric(6, 2), nullable=True)
    stripe_prime_mils = db.Column(db.Numeric(6, 2), nullable=True)
    stripe_prime_gal = db.Column(db.Numeric(10, 2), nullable=True)
    stripe_intermediate_vol_pct = db.Column(db.Numeric(6, 2), nullable=True)
    stripe_intermediate_mils = db.Column(db.Numeric(6, 2), nullable=True)
    stripe_intermediate_gal = db.Column(db.Numeric(10, 2), nullable=True)
    intermediate_vol_pct = db.Column(db.Numeric(6, 2), nullable=True)
    intermediate_mils = db.Column(db.Numeric(6, 2), nullable=True)
    intermediate_gal = db.Column(db.Numeric(10, 2), nullable=True)
    finish_vol_pct = db.Column(db.Numeric(6, 2), nullable=True)
    finish_mils = db.Column(db.Numeric(6, 2), nullable=True)
    finish_gal = db.Column(db.Numeric(10, 2), nullable=True)

    # ── Labor Rates — all tasks: SF/HR + workers/nozzle ───────────────────────
    mobilize_sf_per_hr = db.Column(db.Numeric(10, 2), nullable=True)
    mobilize_workers_per_nozzle = db.Column(db.Numeric(6, 2), nullable=True)
    equip_setup_sf_per_hr = db.Column(db.Numeric(10, 2), nullable=True)
    equip_setup_workers_per_nozzle = db.Column(db.Numeric(6, 2), nullable=True)
    scaffold_sf_per_hr = db.Column(db.Numeric(10, 2), nullable=True)
    scaffold_workers_per_nozzle = db.Column(db.Numeric(6, 2), nullable=True)
    containment_sf_per_hr = db.Column(db.Numeric(10, 2), nullable=True)
    containment_workers_per_nozzle = db.Column(db.Numeric(6, 2), nullable=True)
    masking_sf_per_hr = db.Column(db.Numeric(10, 2), nullable=True)
    masking_workers_per_nozzle = db.Column(db.Numeric(6, 2), nullable=True)
    pressure_wash_sf_per_hr = db.Column(db.Numeric(10, 2), nullable=True)
    pressure_wash_workers_per_nozzle = db.Column(db.Numeric(6, 2), nullable=True)
    caulking_sf_per_hr = db.Column(db.Numeric(10, 2), nullable=True)
    caulking_workers_per_nozzle = db.Column(db.Numeric(6, 2), nullable=True)
    blast_sf_per_hr = db.Column(db.Numeric(10, 2), nullable=True)
    blast_workers_per_nozzle = db.Column(db.Numeric(6, 2), nullable=True)
    primer_labor_sf_per_hr = db.Column(db.Numeric(10, 2), nullable=True)
    primer_labor_workers_per_nozzle = db.Column(db.Numeric(6, 2), nullable=True)
    second_primer_labor_sf_per_hr = db.Column(db.Numeric(10, 2), nullable=True)
    second_primer_labor_workers_per_nozzle = db.Column(db.Numeric(6, 2), nullable=True)
    stripe_prime_labor_sf_per_hr = db.Column(db.Numeric(10, 2), nullable=True)
    stripe_prime_labor_workers_per_nozzle = db.Column(db.Numeric(6, 2), nullable=True)
    stripe_intermediate_labor_sf_per_hr = db.Column(db.Numeric(10, 2), nullable=True)
    stripe_intermediate_labor_workers_per_nozzle = db.Column(db.Numeric(6, 2), nullable=True)
    intermediate_labor_sf_per_hr = db.Column(db.Numeric(10, 2), nullable=True)
    intermediate_labor_workers_per_nozzle = db.Column(db.Numeric(6, 2), nullable=True)
    finish_labor_sf_per_hr = db.Column(db.Numeric(10, 2), nullable=True)
    finish_labor_workers_per_nozzle = db.Column(db.Numeric(6, 2), nullable=True)
    traffic_control_sf_per_hr = db.Column(db.Numeric(10, 2), nullable=True)
    traffic_control_workers_per_nozzle = db.Column(db.Numeric(6, 2), nullable=True)
    inspection_touchup_sf_per_hr = db.Column(db.Numeric(10, 2), nullable=True)
    inspection_touchup_workers_per_nozzle = db.Column(db.Numeric(6, 2), nullable=True)
    osha_training_sf_per_hr = db.Column(db.Numeric(10, 2), nullable=True)
    osha_training_workers_per_nozzle = db.Column(db.Numeric(6, 2), nullable=True)

    # ── Labor Rates — time-based tasks: hrs/day + days ────────────────────────
    mobilize_hrs_per_day = db.Column(db.Numeric(6, 2), nullable=True)
    mobilize_days = db.Column(db.Numeric(6, 2), nullable=True)
    equip_setup_hrs_per_day = db.Column(db.Numeric(6, 2), nullable=True)
    equip_setup_days = db.Column(db.Numeric(6, 2), nullable=True)
    scaffold_hrs_per_day = db.Column(db.Numeric(6, 2), nullable=True)
    scaffold_days = db.Column(db.Numeric(6, 2), nullable=True)
    traffic_control_hrs_per_day = db.Column(db.Numeric(6, 2), nullable=True)
    traffic_control_days = db.Column(db.Numeric(6, 2), nullable=True)
    inspection_touchup_hrs_per_day = db.Column(db.Numeric(6, 2), nullable=True)
    inspection_touchup_days = db.Column(db.Numeric(6, 2), nullable=True)
    osha_training_hrs_per_day = db.Column(db.Numeric(6, 2), nullable=True)
    osha_training_days = db.Column(db.Numeric(6, 2), nullable=True)

    # ── Shift Structure ────────────────────────────────────────────────────────
    shift_hours_per_day = db.Column(db.Numeric(6, 2), nullable=True)
    shift_days_total = db.Column(db.Numeric(8, 2), nullable=True)
    crew_size = db.Column(db.Integer, nullable=True)
    shifts_per_day = db.Column(db.Integer, nullable=True)

    @property
    def scopes_list(self):
        if not self.scopes:
            return []
        try:
            import json
            return json.loads(self.scopes)
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


intelligence_item_tags = db.Table(
    "intelligence_item_tags",
    db.Column("item_id", db.Integer, db.ForeignKey("intelligence_item.id"), primary_key=True),
    db.Column("tag_id", db.Integer, db.ForeignKey("intelligence_tag.id"), primary_key=True),
)


class IntelligenceTag(db.Model):
    __tablename__ = "intelligence_tag"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), unique=True, nullable=False, index=True)
    usage_count = db.Column(db.Integer, nullable=False, default=0)


class IntelligenceItem(db.Model):
    __tablename__ = "intelligence_item"
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(300), nullable=False)
    description = db.Column(db.Text, nullable=True)
    entry_type = db.Column(db.String(10), nullable=False, default="text")  # "text" or "file"
    text_content = db.Column(db.Text, nullable=True)
    file_path = db.Column(db.String(500), nullable=True)
    original_filename = db.Column(db.String(300), nullable=True)
    file_mime = db.Column(db.String(100), nullable=True)
    project_id = db.Column(db.Integer, db.ForeignKey("project.id"), nullable=True)
    work_scope_json = db.Column(db.Text, nullable=True)
    auto_include_in_search = db.Column(db.Boolean, nullable=False, default=True)
    uploaded_by = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    # Quote extraction fields (null for manually-created entries)
    pricing_items_json = db.Column(db.Text, nullable=True)
    conditions_text = db.Column(db.Text, nullable=True)
    flags_json = db.Column(db.Text, nullable=True)
    raw_text_excerpt = db.Column(db.Text, nullable=True)
    extraction_status = db.Column(db.String(20), nullable=True, default="manual")
    vendor_name = db.Column(db.String(200), nullable=True)
    vendor_contact = db.Column(db.Text, nullable=True)
    quote_date = db.Column(db.Date, nullable=True)
    expiration_date = db.Column(db.Date, nullable=True)
    content_hash = db.Column(db.String(64), nullable=True, index=True)
    # Shortlist fields (v1: shortlisted within a category; v2 will populate bid_id/scope_option)
    shortlisted = db.Column(db.Boolean, nullable=False, default=False)
    shortlist_notes = db.Column(db.Text, nullable=True)
    shortlisted_at = db.Column(db.DateTime, nullable=True)
    shortlisted_by = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=True)
    shortlisted_bid_id = db.Column(db.Integer, nullable=True)       # reserved for v2 Bid layer
    shortlisted_scope_option = db.Column(db.String(100), nullable=True)  # reserved for v2

    project = db.relationship("Project", backref=db.backref("intelligence_items", cascade="all, delete-orphan"))
    uploader = db.relationship("User", foreign_keys=[uploaded_by], backref="intelligence_items")
    shortlister = db.relationship("User", foreign_keys=[shortlisted_by], backref="shortlisted_items")
    tags = db.relationship("IntelligenceTag", secondary=intelligence_item_tags, backref="items")

    @property
    def work_scope_list(self):
        if not self.work_scope_json:
            return []
        try:
            import json
            return json.loads(self.work_scope_json)
        except Exception:
            return []


class ComparisonSummary(db.Model):
    __tablename__ = "comparison_summary"
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey("project.id"), nullable=False)
    category_tag = db.Column(db.String(200), nullable=False)
    summary_text = db.Column(db.Text, nullable=False)
    generated_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    skippy_recommendation = db.Column(db.Text, nullable=True)
    takeoff_id = db.Column(db.Integer, db.ForeignKey("takeoff.id"), nullable=True)
    # Reserved for v2 Bid / Scope Option layers
    bid_id = db.Column(db.Integer, nullable=True)
    scope_option = db.Column(db.String(100), nullable=True)

    project = db.relationship("Project", backref=db.backref("comparison_summaries", cascade="all, delete-orphan"))
    takeoff = db.relationship("Takeoff", backref="comparison_summaries")


class QuoteComparisonExport(db.Model):
    __tablename__ = "quote_comparison_export"
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey("project.id"), nullable=False)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    export_type = db.Column(db.String(10), nullable=False)   # "pdf" or "excel"
    category_tag = db.Column(db.String(200), nullable=True)
    vendor_count = db.Column(db.Integer, nullable=True)
    shortlisted_count = db.Column(db.Integer, nullable=True)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    # Reserved for v2 Bid / Scope Option layers
    bid_id = db.Column(db.Integer, nullable=True)
    scope_option = db.Column(db.String(100), nullable=True)

    project = db.relationship("Project", backref=db.backref("quote_comparison_exports", cascade="all, delete-orphan"))
    user = db.relationship("User")


class QuoteBatch(db.Model):
    __tablename__ = "quote_batch"
    id = db.Column(db.Integer, primary_key=True)
    batch_id = db.Column(db.String(36), unique=True, nullable=False, index=True)
    project_id = db.Column(db.Integer, db.ForeignKey("project.id"), nullable=False)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    status = db.Column(db.String(20), nullable=False, default="pending")
    category_tag = db.Column(db.String(100), nullable=True)
    entries_json = db.Column(db.Text, nullable=True)

    takeoff_id = db.Column(db.Integer, db.ForeignKey("takeoff.id"), nullable=True)

    project = db.relationship("Project", backref=db.backref("quote_batches", cascade="all, delete-orphan"))
    user = db.relationship("User")
    takeoff = db.relationship("Takeoff", backref="quote_batches")


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

    takeoff_id = db.Column(db.Integer, db.ForeignKey("takeoff.id"), nullable=True)

    project = db.relationship("Project", backref=db.backref("reports", cascade="all, delete-orphan"))
    user = db.relationship("User")
    takeoff = db.relationship("Takeoff", backref="reports")


class WorkspaceThread(db.Model):
    __tablename__ = "workspace_thread"
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey("project.id"), nullable=False)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    title = db.Column(db.Text, nullable=True)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    updated_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    project = db.relationship("Project", backref=db.backref("workspace_threads", cascade="all, delete-orphan"))
    user = db.relationship("User", backref="workspace_threads")
    messages = db.relationship("WorkspaceMessage", backref="thread", cascade="all, delete-orphan",
                               foreign_keys="WorkspaceMessage.thread_id",
                               order_by="WorkspaceMessage.created_at")

    @property
    def relative_time(self):
        dt = self.updated_at or self.created_at
        if dt is None:
            return ""
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        diff = (datetime.now(timezone.utc) - dt).total_seconds()
        if diff < 60:
            return "just now"
        if diff < 3600:
            return f"{int(diff / 60)}m ago"
        if diff < 86400:
            return f"{int(diff / 3600)}h ago"
        return f"{int(diff / 86400)}d ago"


class WorkspaceMessage(db.Model):
    __tablename__ = "workspace_message"
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey("project.id"), nullable=False)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    thread_id = db.Column(db.Integer, db.ForeignKey("workspace_thread.id"), nullable=True)
    role = db.Column(db.String(20), nullable=False)  # "user" or "assistant"
    content = db.Column(db.Text, nullable=False)
    sources_json = db.Column(db.Text, nullable=True)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    project = db.relationship("Project", backref=db.backref("workspace_messages", cascade="all, delete-orphan"))
    user = db.relationship("User", backref="workspace_messages")
