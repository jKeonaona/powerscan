import io
import os
import uuid
from datetime import datetime, timezone

from flask import (
    Flask, render_template, request, redirect, url_for, flash,
    send_from_directory, send_file, jsonify, abort,
)
from flask_login import (
    LoginManager, login_user, logout_user, login_required, current_user,
)

from config import Config
from models import (
    db, User, Company, Project, Drawing, DrawingPage, SearchHistory, Report, Feedback,
    ROLE_SUPERADMIN, ROLE_ADMIN, ROLE_USER, ROLES,
    DOC_TYPES, DEFAULT_DOC_TYPE,
    FEEDBACK_TYPES, FEEDBACK_STATUSES, DEFAULT_FEEDBACK_STATUS,
)
from email_notify import send_feedback_email_async, send_reply_email_async
from pipeline import start_worker
from reports import REPORT_TEMPLATES, enqueue_report, start_report_worker
from search import search_drawings


CCC_ADMIN_SEEDS = [
    ("orgon@muehlhan.com", "orgon"),
    ("j.brockman@muehlhan.com", "j.brockman"),
    ("lasater@muehlhan.com", "lasater"),
    ("moore@muehlhan.com", "moore"),
]
CCC_ADMIN_TEMP_PASSWORD = "Temp?Access123"
CCC_COMPANY_ID = 1


def _seed_ccc_admins():
    """Idempotently ensure the four CCC admin accounts exist with a forced password reset.

    Skipped silently if the target company does not yet exist, and per-user if an
    account with that email already exists (we do NOT reset an existing user's
    password here — that would be a footgun for subsequent deploys).
    """
    company = db.session.get(Company, CCC_COMPANY_ID)
    if not company:
        print(f"[powerscan] CCC seed: company id={CCC_COMPANY_ID} not found, skipping admin seed", flush=True)
        return

    created = 0
    for email, username in CCC_ADMIN_SEEDS:
        if User.query.filter_by(email=email).first():
            continue
        if User.query.filter_by(username=username).first():
            print(f"[powerscan] CCC seed: username '{username}' taken, skipping {email}", flush=True)
            continue
        user = User(
            username=username,
            email=email,
            role=ROLE_ADMIN,
            company_id=CCC_COMPANY_ID,
            must_change_password=True,
        )
        user.set_password(CCC_ADMIN_TEMP_PASSWORD)
        db.session.add(user)
        created += 1
    if created:
        db.session.commit()
        print(f"[powerscan] CCC seed: created {created} admin account(s)", flush=True)


def _run_migrations(database):
    """Add any missing columns to existing tables via ALTER TABLE."""
    conn = database.engine.raw_connection()
    cursor = conn.cursor()
    # Each entry: (table, column, column_def)
    migrations = [
        ("drawing", "total_pages", "INTEGER DEFAULT 0"),
        ("drawing", "pages_processed", "INTEGER DEFAULT 0"),
        ("drawing", "doc_type", "VARCHAR(40) DEFAULT 'Drawing'"),
        ("report", "file_path", "VARCHAR(300)"),
        ("user", "must_change_password", "BOOLEAN DEFAULT 0 NOT NULL"),
        ("feedback", "admin_reply", "TEXT"),
    ]
    for table, column, col_def in migrations:
        try:
            cursor.execute(f"ALTER TABLE {table} ADD COLUMN {column} {col_def}")
        except Exception:
            pass  # Column already exists
    conn.commit()
    conn.close()


def create_app():
    app = Flask(__name__)
    app.config.from_object(Config)

    os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
    os.makedirs(app.config["PROCESSED_FOLDER"], exist_ok=True)
    os.makedirs(app.config["REPORTS_FOLDER"], exist_ok=True)
    os.makedirs(os.path.join(app.instance_path), exist_ok=True)

    db.init_app(app)

    login_manager = LoginManager()
    login_manager.login_view = "login"
    login_manager.init_app(app)

    @login_manager.user_loader
    def load_user(user_id):
        return db.session.get(User, int(user_id))

    with app.app_context():
        db.create_all()
        # Migrate: add columns that may be missing from older databases
        _run_migrations(db)
        # Create default superadmin if none exists
        if not User.query.filter_by(role=ROLE_SUPERADMIN).first():
            admin = User(
                username="admin",
                email="admin@powerscan.local",
                role=ROLE_SUPERADMIN,
            )
            admin.set_password("admin123")
            db.session.add(admin)
            db.session.commit()

        # Seed CCC admin accounts (idempotent — skipped if email already exists)
        _seed_ccc_admins()

    # Start background conversion worker thread
    start_worker(app)
    start_report_worker(app)

    # ── Decorators ──────────────────────────────────────────────

    def admin_required(f):
        from functools import wraps

        @wraps(f)
        def decorated(*args, **kwargs):
            if not current_user.is_authenticated or not current_user.is_admin:
                abort(403)
            return f(*args, **kwargs)
        return decorated

    def superadmin_required(f):
        from functools import wraps

        @wraps(f)
        def decorated(*args, **kwargs):
            if not current_user.is_authenticated or not current_user.is_superadmin:
                abort(403)
            return f(*args, **kwargs)
        return decorated

    # ── Auth Routes ─────────────────────────────────────────────

    @app.route("/login", methods=["GET", "POST"])
    def login():
        if current_user.is_authenticated:
            if current_user.must_change_password:
                return redirect(url_for("change_password"))
            return redirect(url_for("dashboard"))
        if request.method == "POST":
            identifier = request.form.get("username", "").strip()
            password = request.form.get("password", "")
            user = (
                User.query.filter_by(username=identifier).first()
                or User.query.filter_by(email=identifier).first()
            )
            if user and user.check_password(password):
                login_user(user)
                if user.must_change_password:
                    flash("Please choose a new password to finish logging in.", "info")
                    return redirect(url_for("change_password"))
                next_page = request.args.get("next")
                return redirect(next_page or url_for("dashboard"))
            flash("Invalid username or password.", "danger")
        return render_template("login.html")

    @app.route("/change-password", methods=["GET", "POST"])
    @login_required
    def change_password():
        if request.method == "POST":
            new_password = request.form.get("new_password", "")
            confirm_password = request.form.get("confirm_password", "")
            if len(new_password) < 8:
                flash("Password must be at least 8 characters.", "danger")
            elif new_password != confirm_password:
                flash("Passwords do not match.", "danger")
            elif current_user.check_password(new_password):
                flash("New password must be different from your current password.", "danger")
            else:
                current_user.set_password(new_password)
                current_user.must_change_password = False
                db.session.commit()
                flash("Password updated. Welcome!", "success")
                return redirect(url_for("dashboard"))
        return render_template("change_password.html", forced=current_user.must_change_password)

    # Guard: users with a forced reset flag are locked to the change-password page
    # until they pick a new password (or log out).
    @app.before_request
    def _enforce_password_reset():
        if not current_user.is_authenticated:
            return None
        if not current_user.must_change_password:
            return None
        allowed_endpoints = {"change_password", "logout", "login", "static"}
        if request.endpoint in allowed_endpoints:
            return None
        return redirect(url_for("change_password"))

    @app.route("/logout")
    @login_required
    def logout():
        logout_user()
        return redirect(url_for("login"))

    # ── Dashboard ───────────────────────────────────────────────

    @app.route("/")
    @login_required
    def dashboard():
        if current_user.is_superadmin:
            companies = Company.query.all()
        elif current_user.company_id:
            companies = [current_user.company]
        else:
            companies = []

        stats = {
            "companies": Company.query.count() if current_user.is_superadmin else len(companies),
            "projects": sum(len(c.projects) for c in companies),
            "drawings": sum(
                len(p.drawings) for c in companies for p in c.projects
            ),
        }
        return render_template("dashboard.html", companies=companies, stats=stats)

    # ── Company Routes ──────────────────────────────────────────

    @app.route("/companies")
    @login_required
    def companies():
        if current_user.is_superadmin:
            company_list = Company.query.order_by(Company.name).all()
        elif current_user.company_id:
            company_list = [current_user.company]
        else:
            company_list = []
        return render_template("companies.html", companies=company_list)

    @app.route("/companies/new", methods=["GET", "POST"])
    @login_required
    @admin_required
    def new_company():
        if request.method == "POST":
            name = request.form.get("name", "").strip()
            if not name:
                flash("Company name is required.", "danger")
            elif Company.query.filter_by(name=name).first():
                flash("Company already exists.", "danger")
            else:
                company = Company(name=name)
                db.session.add(company)
                db.session.commit()
                flash(f"Company '{name}' created.", "success")
                return redirect(url_for("companies"))
        return render_template("company_form.html")

    @app.route("/companies/<int:company_id>/delete", methods=["POST"])
    @login_required
    @superadmin_required
    def delete_company(company_id):
        company = db.session.get(Company, company_id) or abort(404)
        db.session.delete(company)
        db.session.commit()
        flash(f"Company '{company.name}' deleted.", "success")
        return redirect(url_for("companies"))

    # ── Project Routes ──────────────────────────────────────────

    @app.route("/companies/<int:company_id>/projects")
    @login_required
    def projects(company_id):
        company = db.session.get(Company, company_id) or abort(404)
        if not current_user.is_superadmin and current_user.company_id != company_id:
            abort(403)
        return render_template("projects.html", company=company)

    @app.route("/companies/<int:company_id>/projects/new", methods=["GET", "POST"])
    @login_required
    @admin_required
    def new_project(company_id):
        company = db.session.get(Company, company_id) or abort(404)
        if request.method == "POST":
            name = request.form.get("name", "").strip()
            description = request.form.get("description", "").strip()
            if not name:
                flash("Project name is required.", "danger")
            else:
                project = Project(name=name, description=description, company_id=company.id)
                db.session.add(project)
                db.session.commit()
                flash(f"Project '{name}' created.", "success")
                return redirect(url_for("projects", company_id=company.id))
        return render_template("project_form.html", company=company)

    @app.route("/projects/<int:project_id>/delete", methods=["POST"])
    @login_required
    @admin_required
    def delete_project(project_id):
        project = db.session.get(Project, project_id) or abort(404)
        company_id = project.company_id
        db.session.delete(project)
        db.session.commit()
        flash(f"Project '{project.name}' deleted.", "success")
        return redirect(url_for("projects", company_id=company_id))

    # ── Drawing Routes ──────────────────────────────────────────

    @app.route("/projects/<int:project_id>/drawings", methods=["GET", "POST"])
    @login_required
    def drawings(project_id):
        project = db.session.get(Project, project_id) or abort(404)
        if not current_user.is_superadmin and current_user.company_id != project.company_id:
            abort(403)

        filter_doc_type = request.values.get("filter_doc_type", "").strip()
        if filter_doc_type and filter_doc_type not in DOC_TYPES:
            filter_doc_type = ""

        results = None
        query = ""
        search_doc_type = ""
        if request.method == "POST":
            query = request.form.get("query", "").strip()
            search_doc_type = request.form.get("search_doc_type", "").strip()
            if search_doc_type and search_doc_type not in DOC_TYPES:
                search_doc_type = ""
            if not query:
                flash("Please enter a question.", "danger")
            elif not app.config["ANTHROPIC_API_KEY"]:
                flash("Claude API key not configured. Set ANTHROPIC_API_KEY environment variable.", "danger")
            else:
                results = search_drawings(
                    query,
                    project.id,
                    app.config["ANTHROPIC_API_KEY"],
                    app.config["PROCESSED_FOLDER"],
                    doc_type=search_doc_type or None,
                )
                history = SearchHistory(
                    project_id=project.id,
                    user_id=current_user.id,
                    query=query,
                    answer=results.get("answer", "") if results else "",
                    doc_type_filter=search_doc_type or None,
                )
                db.session.add(history)
                db.session.commit()

        drawings_q = Drawing.query.filter_by(project_id=project.id)
        if filter_doc_type:
            drawings_q = drawings_q.filter_by(doc_type=filter_doc_type)
        drawings_list = drawings_q.order_by(Drawing.created_at.desc()).all()

        reports_list = (
            Report.query.filter_by(project_id=project.id)
            .order_by(Report.created_at.desc())
            .limit(20)
            .all()
        )

        history_entries = (
            db.session.query(SearchHistory)
            .filter_by(project_id=project.id)
            .order_by(SearchHistory.created_at.desc())
            .limit(50)
            .all()
        )

        active_tab = request.args.get("tab", "").strip().lower()
        if active_tab not in ("documents", "search", "reports", "history"):
            if request.method == "POST" or query:
                active_tab = "search"
            else:
                active_tab = "documents"

        has_ready_docs = Drawing.query.filter_by(project_id=project.id, status="ready").count() > 0

        return render_template(
            "drawings.html",
            project=project,
            drawings_list=drawings_list,
            results=results,
            query=query,
            doc_types=DOC_TYPES,
            filter_doc_type=filter_doc_type,
            search_doc_type=search_doc_type,
            report_templates=REPORT_TEMPLATES,
            reports_list=reports_list,
            history_entries=history_entries,
            active_tab=active_tab,
            has_ready_docs=has_ready_docs,
        )

    @app.route("/projects/<int:project_id>/upload", methods=["GET", "POST"])
    @login_required
    def upload_drawing(project_id):
        project = db.session.get(Project, project_id) or abort(404)
        if not current_user.is_superadmin and current_user.company_id != project.company_id:
            abort(403)
        return render_template("upload.html", project=project, doc_types=DOC_TYPES)

    @app.route("/projects/<int:project_id>/upload-file", methods=["POST"])
    @login_required
    def upload_single_file(project_id):
        """AJAX endpoint: accepts one PDF at a time, returns JSON."""
        project = db.session.get(Project, project_id) or abort(404)
        if not current_user.is_superadmin and current_user.company_id != project.company_id:
            abort(403)

        file = request.files.get("pdf_file")
        if not file or not file.filename.lower().endswith(".pdf"):
            return jsonify({"error": "Invalid file. Only PDFs are accepted."}), 400

        replace = request.form.get("replace") == "1"
        doc_type = request.form.get("doc_type", DEFAULT_DOC_TYPE)
        if doc_type not in DOC_TYPES:
            doc_type = DEFAULT_DOC_TYPE

        # Check for duplicate filename in this project
        existing = Drawing.query.filter_by(
            project_id=project.id,
            original_filename=file.filename,
        ).first()

        if existing and not replace:
            return jsonify({
                "duplicate": True,
                "filename": file.filename,
                "existing_id": existing.id,
                "existing_status": existing.status,
            }), 409

        # If replacing, delete the old drawing
        if existing and replace:
            db.session.delete(existing)
            db.session.commit()

        ext = os.path.splitext(file.filename)[1]
        unique_name = f"{uuid.uuid4().hex}{ext}"
        file.save(os.path.join(app.config["UPLOAD_FOLDER"], unique_name))

        drawing = Drawing(
            filename=unique_name,
            original_filename=file.filename,
            project_id=project.id,
            uploaded_by=current_user.id,
            doc_type=doc_type,
            status="pending",
        )
        db.session.add(drawing)
        db.session.commit()

        return jsonify({
            "id": drawing.id,
            "filename": file.filename,
            "status": "pending",
        })

    @app.route("/drawings/<int:drawing_id>")
    @login_required
    def drawing_detail(drawing_id):
        drawing = db.session.get(Drawing, drawing_id) or abort(404)
        if not current_user.is_superadmin and current_user.company_id != drawing.project.company_id:
            abort(403)
        pages = DrawingPage.query.filter_by(drawing_id=drawing.id).order_by(DrawingPage.page_number).all()
        return render_template("drawing_detail.html", drawing=drawing, pages=pages)

    @app.route("/drawings/<int:drawing_id>/status")
    @login_required
    def drawing_status(drawing_id):
        drawing = db.session.get(Drawing, drawing_id) or abort(404)
        return jsonify({
            "status": drawing.status,
            "total_pages": drawing.total_pages,
            "pages_processed": drawing.pages_processed,
        })

    @app.route("/drawings/<int:drawing_id>/reprocess", methods=["POST"])
    @login_required
    def reprocess_drawing(drawing_id):
        drawing = db.session.get(Drawing, drawing_id) or abort(404)
        if not current_user.is_superadmin and current_user.company_id != drawing.project.company_id:
            abort(403)
        if drawing.status == "processing":
            flash("Document is already being processed.", "warning")
            return redirect(url_for("drawing_detail", drawing_id=drawing.id))

        drawing.status = "pending"
        drawing.pages_processed = 0
        db.session.commit()
        flash("Reconverting PDF pages to images.", "success")
        return redirect(url_for("drawing_detail", drawing_id=drawing.id))

    @app.route("/drawings/<int:drawing_id>/delete", methods=["POST"])
    @login_required
    @admin_required
    def delete_drawing(drawing_id):
        drawing = db.session.get(Drawing, drawing_id) or abort(404)
        project_id = drawing.project_id
        db.session.delete(drawing)
        db.session.commit()
        flash("Document deleted.", "success")
        return redirect(url_for("drawings", project_id=project_id))

    @app.route("/processed/<path:filename>")
    @login_required
    def serve_processed(filename):
        return send_from_directory(app.config["PROCESSED_FOLDER"], filename)

    # ── Search History ──────────────────────────────────────────

    def _project_history(project_id):
        project = db.session.get(Project, project_id) or abort(404)
        if not current_user.is_superadmin and current_user.company_id != project.company_id:
            abort(403)
        entries = (
            db.session.query(SearchHistory)
            .filter_by(project_id=project.id)
            .order_by(SearchHistory.created_at.desc())
            .all()
        )
        return project, entries

    @app.route("/projects/<int:project_id>/history")
    @login_required
    def search_history(project_id):
        project, entries = _project_history(project_id)
        return render_template("search_history.html", project=project, entries=entries)

    @app.route("/projects/<int:project_id>/report/generate", methods=["POST"])
    @login_required
    def generate_project_report(project_id):
        project = db.session.get(Project, project_id) or abort(404)
        if not current_user.is_superadmin and current_user.company_id != project.company_id:
            abort(403)

        if not app.config["ANTHROPIC_API_KEY"]:
            flash("Claude API key not configured.", "danger")
            return redirect(url_for("drawings", project_id=project.id))

        template_id = request.form.get("template_id", "").strip()
        custom_prompt = request.form.get("custom_prompt", "").strip()
        if not template_id or (template_id != "custom" and template_id not in REPORT_TEMPLATES):
            flash("Please choose a valid report template.", "danger")
            return redirect(url_for("drawings", project_id=project.id))

        try:
            enqueue_report(project.id, current_user.id, template_id, custom_prompt)
        except ValueError as e:
            flash(str(e), "danger")
            return redirect(url_for("drawings", project_id=project.id))

        flash("Report is being generated in the background. It will appear in the Reports tab when ready.", "info")
        return redirect(url_for("drawings", project_id=project.id, tab="reports"))

    @app.route("/reports/<int:report_id>/download")
    @login_required
    def download_report(report_id):
        report = db.session.get(Report, report_id) or abort(404)
        project = db.session.get(Project, report.project_id) or abort(404)
        if not current_user.is_superadmin and current_user.company_id != project.company_id:
            abort(403)
        if report.status != "ready" or not report.file_path:
            flash("Report is not ready yet.", "warning")
            return redirect(url_for("drawings", project_id=project.id, tab="reports"))
        file_path = os.path.join(app.config["REPORTS_FOLDER"], report.file_path)
        with open(file_path, "rb") as f:
            data = f.read()
        return send_file(
            io.BytesIO(data),
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            as_attachment=True,
            download_name=report.file_path,
        )

    @app.route("/reports/<int:report_id>/download/pdf")
    @login_required
    def download_report_pdf(report_id):
        import subprocess
        import shutil
        import tempfile

        report = db.session.get(Report, report_id) or abort(404)
        project = db.session.get(Project, report.project_id) or abort(404)
        if not current_user.is_superadmin and current_user.company_id != project.company_id:
            abort(403)
        if report.status != "ready" or not report.file_path:
            flash("Report is not ready yet.", "warning")
            return redirect(url_for("drawings", project_id=project.id, tab="reports"))

        docx_path = os.path.join(app.config["REPORTS_FOLDER"], report.file_path)
        pdf_name = os.path.splitext(report.file_path)[0] + ".pdf"

        # Try LibreOffice conversion first
        pdf_bytes = None
        tmp_dir = tempfile.mkdtemp()
        try:
            result = subprocess.run(
                ["libreoffice", "--headless", "--convert-to", "pdf",
                 "--outdir", tmp_dir, docx_path],
                capture_output=True,
                timeout=60,
            )
            if result.returncode == 0:
                base = os.path.splitext(os.path.basename(docx_path))[0]
                pdf_path = os.path.join(tmp_dir, base + ".pdf")
                if os.path.exists(pdf_path):
                    with open(pdf_path, "rb") as f:
                        pdf_bytes = f.read()
        except Exception:
            pass
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)

        # Fall back to reportlab if LibreOffice unavailable or failed
        if pdf_bytes is None:
            from docx import Document as DocxDocument
            from reportlab.lib.pagesizes import letter
            from reportlab.lib.styles import getSampleStyleSheet
            from reportlab.lib.units import inch
            from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
            from xml.sax.saxutils import escape as xml_escape

            buf = io.BytesIO()
            rdoc = SimpleDocTemplate(
                buf, pagesize=letter,
                leftMargin=0.75 * inch, rightMargin=0.75 * inch,
                topMargin=0.75 * inch, bottomMargin=0.75 * inch,
            )
            styles = getSampleStyleSheet()
            story = []
            try:
                docx = DocxDocument(docx_path)
                for para in docx.paragraphs:
                    text = para.text.strip()
                    if not text:
                        story.append(Spacer(1, 6))
                        continue
                    style_name = para.style.name if para.style else "Normal"
                    if "Heading 1" in style_name:
                        story.append(Paragraph(xml_escape(text), styles["Heading1"]))
                    elif "Heading 2" in style_name:
                        story.append(Paragraph(xml_escape(text), styles["Heading2"]))
                    elif "Heading 3" in style_name:
                        story.append(Paragraph(xml_escape(text), styles["Heading3"]))
                    else:
                        story.append(Paragraph(xml_escape(text), styles["BodyText"]))
                for table in docx.tables:
                    for row in table.rows:
                        cells = " | ".join(xml_escape(c.text.strip()) for c in row.cells)
                        story.append(Paragraph(cells, styles["Normal"]))
                    story.append(Spacer(1, 6))
            except Exception:
                story.append(Paragraph("(could not read report content)", styles["Normal"]))
            if not story:
                story.append(Paragraph("(empty report)", styles["Normal"]))
            rdoc.build(story)
            buf.seek(0)
            pdf_bytes = buf.read()

        return send_file(
            io.BytesIO(pdf_bytes),
            mimetype="application/pdf",
            as_attachment=True,
            download_name=pdf_name,
        )

    @app.route("/reports/<int:report_id>/status")
    @login_required
    def report_status(report_id):
        report = db.session.get(Report, report_id) or abort(404)
        project = db.session.get(Project, report.project_id) or abort(404)
        if not current_user.is_superadmin and current_user.company_id != project.company_id:
            abort(403)
        return jsonify({
            "id": report.id,
            "status": report.status,
            "file_path": report.file_path,
            "error_message": report.error_message,
        })

    @app.route("/reports/<int:report_id>/delete", methods=["POST"])
    @login_required
    def delete_report(report_id):
        report = db.session.get(Report, report_id) or abort(404)
        project = db.session.get(Project, report.project_id) or abort(404)
        if not current_user.is_superadmin and current_user.company_id != project.company_id:
            abort(403)
        if report.file_path:
            path = os.path.join(app.config["REPORTS_FOLDER"], report.file_path)
            try:
                os.remove(path)
            except OSError:
                pass
        db.session.delete(report)
        db.session.commit()
        flash("Report deleted.", "success")
        return redirect(url_for("drawings", project_id=project.id, tab="reports"))

    @app.route("/projects/<int:project_id>/history/export")
    @login_required
    def export_search_history(project_id):
        from reportlab.lib.pagesizes import letter
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import inch
        from reportlab.platypus import (
            SimpleDocTemplate, Paragraph, Spacer, PageBreak, HRFlowable,
        )
        from reportlab.lib import colors
        from xml.sax.saxutils import escape as xml_escape

        project, entries = _project_history(project_id)

        buf = io.BytesIO()
        doc = SimpleDocTemplate(
            buf, pagesize=letter,
            leftMargin=0.75 * inch, rightMargin=0.75 * inch,
            topMargin=0.75 * inch, bottomMargin=0.75 * inch,
            title=f"Search History — {project.name}",
        )
        styles = getSampleStyleSheet()
        h_title = styles["Title"]
        h_meta = ParagraphStyle("meta", parent=styles["Normal"], textColor=colors.grey, fontSize=9, spaceAfter=12)
        h_q = ParagraphStyle("q", parent=styles["Heading3"], textColor=colors.HexColor("#0d6efd"), spaceAfter=4)
        h_entry_meta = ParagraphStyle("em", parent=styles["Normal"], textColor=colors.grey, fontSize=8, spaceAfter=6)
        h_a = ParagraphStyle("a", parent=styles["BodyText"], leading=14, spaceAfter=10)

        def p(text, style):
            return Paragraph(xml_escape(text or "").replace("\n", "<br/>"), style)

        story = [
            Paragraph(f"Search History — {xml_escape(project.name)}", h_title),
            Paragraph(
                f"{xml_escape(project.company.name)} &middot; {len(entries)} search(es) &middot; "
                f"Generated {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M UTC')}",
                h_meta,
            ),
            HRFlowable(width="100%", thickness=0.5, color=colors.lightgrey, spaceAfter=12),
        ]

        if not entries:
            story.append(Paragraph("No searches recorded for this project.", styles["Italic"]))
        else:
            for i, e in enumerate(entries, start=1):
                who = e.user.username if e.user else "unknown"
                when = e.created_at.strftime("%Y-%m-%d %H:%M UTC")
                filt = f" &middot; filter: {xml_escape(e.doc_type_filter)}" if e.doc_type_filter else ""
                story.append(p(f"{i}. {e.query}", h_q))
                story.append(Paragraph(f"{xml_escape(who)} &middot; {when}{filt}", h_entry_meta))
                story.append(p(e.answer or "(no answer recorded)", h_a))
                if i < len(entries):
                    story.append(HRFlowable(width="100%", thickness=0.3, color=colors.whitesmoke, spaceAfter=10))

        doc.build(story)
        buf.seek(0)
        fname = f"search-history-{project.id}-{datetime.now(timezone.utc).strftime('%Y%m%d-%H%M')}.pdf"
        return send_file(
            buf,
            mimetype="application/pdf",
            as_attachment=True,
            download_name=fname,
        )

    # ── Search ──────────────────────────────────────────────────

    @app.route("/search", methods=["GET", "POST"])
    @login_required
    def search():
        if not current_user.is_superadmin:
            abort(403)
        projects_list = (
            Project.query.join(Company).order_by(Company.name, Project.name).all()
        )

        results = None
        query = ""
        selected_project_id = None
        search_doc_type = ""
        if request.method == "POST":
            query = request.form.get("query", "").strip()
            selected_project_id = request.form.get("project_id", type=int)
            search_doc_type = request.form.get("search_doc_type", "").strip()
            if search_doc_type and search_doc_type not in DOC_TYPES:
                search_doc_type = ""
            if not query:
                flash("Please enter a search query.", "danger")
            elif not selected_project_id:
                flash("Please select a project.", "danger")
            elif not app.config["ANTHROPIC_API_KEY"]:
                flash("Claude API key not configured. Set ANTHROPIC_API_KEY environment variable.", "danger")
            else:
                project = db.session.get(Project, selected_project_id)
                if not project:
                    flash("Project not found.", "danger")
                elif not current_user.is_superadmin and current_user.company_id != project.company_id:
                    abort(403)
                else:
                    results = search_drawings(
                        query,
                        selected_project_id,
                        app.config["ANTHROPIC_API_KEY"],
                        app.config["PROCESSED_FOLDER"],
                        doc_type=search_doc_type or None,
                    )
                    history = SearchHistory(
                        project_id=selected_project_id,
                        user_id=current_user.id,
                        query=query,
                        answer=results.get("answer", "") if results else "",
                        doc_type_filter=search_doc_type or None,
                    )
                    db.session.add(history)
                    db.session.commit()

        return render_template(
            "search.html",
            results=results,
            query=query,
            projects=projects_list,
            selected_project_id=selected_project_id,
            doc_types=DOC_TYPES,
            search_doc_type=search_doc_type,
        )

    # ── Admin: User Management ──────────────────────────────────

    @app.route("/admin/users")
    @login_required
    @admin_required
    def admin_users():
        if current_user.is_superadmin:
            users = User.query.order_by(User.username).all()
        else:
            users = User.query.filter_by(company_id=current_user.company_id).order_by(User.username).all()
        companies = Company.query.order_by(Company.name).all()
        return render_template("admin/users.html", users=users, companies=companies, roles=ROLES)

    @app.route("/admin/users/new", methods=["GET", "POST"])
    @login_required
    @admin_required
    def new_user():
        companies = Company.query.order_by(Company.name).all()
        if request.method == "POST":
            username = request.form.get("username", "").strip()
            email = request.form.get("email", "").strip()
            password = request.form.get("password", "")
            role = request.form.get("role", ROLE_USER)
            company_id = request.form.get("company_id", type=int)

            if not all([username, email, password]):
                flash("All fields are required.", "danger")
            elif User.query.filter_by(username=username).first():
                flash("Username already exists.", "danger")
            elif User.query.filter_by(email=email).first():
                flash("Email already exists.", "danger")
            else:
                if not current_user.is_superadmin:
                    role = ROLE_USER
                    company_id = current_user.company_id

                user = User(username=username, email=email, role=role, company_id=company_id)
                user.set_password(password)
                db.session.add(user)
                db.session.commit()
                flash(f"User '{username}' created.", "success")
                return redirect(url_for("admin_users"))

        allowed_roles = ROLES if current_user.is_superadmin else [ROLE_USER]
        return render_template("admin/user_form.html", companies=companies, roles=allowed_roles)

    @app.route("/admin/users/<int:user_id>/edit", methods=["GET", "POST"])
    @login_required
    @admin_required
    def edit_user(user_id):
        user = db.session.get(User, user_id) or abort(404)
        companies = Company.query.order_by(Company.name).all()

        if request.method == "POST":
            user.email = request.form.get("email", user.email).strip()
            new_password = request.form.get("password", "").strip()
            if new_password:
                user.set_password(new_password)
            if current_user.is_superadmin:
                user.role = request.form.get("role", user.role)
                user.company_id = request.form.get("company_id", type=int)
            db.session.commit()
            flash(f"User '{user.username}' updated.", "success")
            return redirect(url_for("admin_users"))

        allowed_roles = ROLES if current_user.is_superadmin else [ROLE_USER]
        return render_template("admin/user_form.html", user=user, companies=companies, roles=allowed_roles)

    @app.route("/admin/users/<int:user_id>/delete", methods=["POST"])
    @login_required
    @superadmin_required
    def delete_user(user_id):
        user = db.session.get(User, user_id) or abort(404)
        if user.id == current_user.id:
            flash("You cannot delete yourself.", "danger")
            return redirect(url_for("admin_users"))
        db.session.delete(user)
        db.session.commit()
        flash(f"User '{user.username}' deleted.", "success")
        return redirect(url_for("admin_users"))

    # ── Feedback ────────────────────────────────────────────────

    @app.route("/feedback/new", methods=["POST"])
    @login_required
    def submit_feedback():
        fb_type = (request.form.get("type") or "").strip()
        page = (request.form.get("page") or "").strip()[:500] or None
        description = (request.form.get("description") or "").strip()

        if fb_type not in FEEDBACK_TYPES:
            return jsonify({"ok": False, "error": "Please choose a feedback type."}), 400
        if not description:
            return jsonify({"ok": False, "error": "Description is required."}), 400
        if len(description) > 10000:
            return jsonify({"ok": False, "error": "Description is too long (max 10,000 characters)."}), 400

        entry = Feedback(
            user_id=current_user.id,
            type=fb_type,
            page=page,
            description=description,
            status=DEFAULT_FEEDBACK_STATUS,
        )
        db.session.add(entry)
        db.session.commit()

        # Fire-and-forget email notification — failures are logged but never block submission.
        send_feedback_email_async(app, entry.id)

        return jsonify({"ok": True, "id": entry.id})

    @app.route("/admin/feedback")
    @login_required
    @admin_required
    def admin_feedback():
        filter_type = (request.args.get("type") or "").strip()
        q = db.session.query(Feedback)
        if filter_type in FEEDBACK_TYPES:
            q = q.filter(Feedback.type == filter_type)
        entries = q.order_by(Feedback.created_at.desc()).all()

        counts = {
            "All": db.session.query(Feedback).count(),
        }
        for t in FEEDBACK_TYPES:
            counts[t] = db.session.query(Feedback).filter_by(type=t).count()

        return render_template(
            "admin/feedback.html",
            entries=entries,
            filter_type=filter_type,
            feedback_types=FEEDBACK_TYPES,
            feedback_statuses=FEEDBACK_STATUSES,
            counts=counts,
        )

    @app.route("/admin/feedback/<int:feedback_id>/update", methods=["POST"])
    @login_required
    @superadmin_required
    def update_feedback(feedback_id):
        entry = db.session.get(Feedback, feedback_id) or abort(404)
        new_status = (request.form.get("status") or "").strip()
        if new_status not in FEEDBACK_STATUSES:
            flash("Invalid status.", "danger")
            return redirect(url_for("admin_feedback"))
        entry.status = new_status
        entry.admin_notes = (request.form.get("admin_notes") or "").strip() or None
        db.session.commit()
        flash("Feedback updated.", "success")
        return redirect(url_for("admin_feedback", type=request.args.get("type", "")))

    @app.route("/admin/feedback/<int:feedback_id>/reply", methods=["POST"])
    @login_required
    @superadmin_required
    def reply_feedback(feedback_id):
        entry = db.session.get(Feedback, feedback_id) or abort(404)
        reply_text = (request.form.get("admin_reply") or "").strip()
        if not reply_text:
            flash("Reply cannot be empty.", "danger")
            return redirect(url_for("admin_feedback", type=request.args.get("type", "")))
        entry.admin_reply = reply_text
        entry.status = "Reviewed"
        db.session.commit()
        send_reply_email_async(app, feedback_id)
        flash("Reply sent and status set to Reviewed.", "success")
        return redirect(url_for("admin_feedback", type=request.args.get("type", "")))

    @app.route("/admin/feedback/export")
    @login_required
    @admin_required
    def export_feedback():
        from reportlab.lib.pagesizes import letter
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import inch
        from reportlab.platypus import (
            SimpleDocTemplate, Paragraph, Spacer, HRFlowable,
        )
        from reportlab.lib import colors
        from xml.sax.saxutils import escape as xml_escape

        entries = (
            db.session.query(Feedback)
            .order_by(Feedback.type, Feedback.created_at.desc())
            .all()
        )

        grouped = {t: [] for t in FEEDBACK_TYPES}
        for e in entries:
            grouped.setdefault(e.type, []).append(e)

        buf = io.BytesIO()
        doc = SimpleDocTemplate(
            buf, pagesize=letter,
            leftMargin=0.75 * inch, rightMargin=0.75 * inch,
            topMargin=0.75 * inch, bottomMargin=0.75 * inch,
            title="PowerScan — User Feedback & Feature Ideas",
        )
        styles = getSampleStyleSheet()
        h_title = styles["Title"]
        h_meta = ParagraphStyle("meta", parent=styles["Normal"], textColor=colors.grey, fontSize=9, spaceAfter=12)
        h_group = ParagraphStyle(
            "group", parent=styles["Heading1"],
            textColor=colors.HexColor("#0d6efd"), fontSize=16, spaceBefore=14, spaceAfter=6,
        )
        h_entry_q = ParagraphStyle(
            "eq", parent=styles["Heading3"],
            textColor=colors.HexColor("#212529"), fontSize=11, spaceAfter=3,
        )
        h_entry_meta = ParagraphStyle(
            "em", parent=styles["Normal"], textColor=colors.grey, fontSize=8, spaceAfter=4,
        )
        h_body = ParagraphStyle("b", parent=styles["BodyText"], leading=13, spaceAfter=4)
        h_notes = ParagraphStyle(
            "n", parent=styles["BodyText"], leading=13, spaceAfter=8,
            textColor=colors.HexColor("#555555"), leftIndent=12,
        )

        def p(text, style):
            return Paragraph(xml_escape(text or "").replace("\n", "<br/>"), style)

        story = [
            Paragraph("PowerScan — User Feedback &amp; Feature Ideas", h_title),
            Paragraph(
                f"Exported {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M UTC')} &middot; "
                f"{len(entries)} total submission(s)",
                h_meta,
            ),
            HRFlowable(width="100%", thickness=0.5, color=colors.lightgrey, spaceAfter=10),
        ]

        if not entries:
            story.append(Paragraph("No feedback has been submitted yet.", styles["Italic"]))
        else:
            for group_type in FEEDBACK_TYPES:
                items = grouped.get(group_type) or []
                if not items:
                    continue
                story.append(Paragraph(f"{group_type} ({len(items)})", h_group))
                story.append(HRFlowable(width="100%", thickness=0.3, color=colors.HexColor("#cfe2ff"), spaceAfter=6))
                for i, e in enumerate(items, start=1):
                    who = e.user.username if e.user else "unknown"
                    when = e.created_at.strftime("%Y-%m-%d %H:%M UTC")
                    status_line = f"Status: {e.status}"
                    if e.page:
                        status_line += f" &middot; Page: {xml_escape(e.page)}"
                    story.append(p(f"{i}. {e.description[:120]}{'…' if len(e.description) > 120 else ''}", h_entry_q))
                    story.append(Paragraph(f"{xml_escape(who)} &middot; {when} &middot; {status_line}", h_entry_meta))
                    story.append(p(e.description, h_body))
                    if e.admin_notes:
                        story.append(p(f"Admin notes: {e.admin_notes}", h_notes))
                    story.append(Spacer(1, 4))
                story.append(Spacer(1, 8))

        doc.build(story)
        buf.seek(0)
        fname = f"powerscan-feedback-{datetime.now(timezone.utc).strftime('%Y%m%d-%H%M')}.pdf"
        return send_file(
            buf,
            mimetype="application/pdf",
            as_attachment=True,
            download_name=fname,
        )

    # ── Error Handlers ──────────────────────────────────────────

    @app.errorhandler(403)
    def forbidden(e):
        return render_template("error.html", code=403, message="Access Denied"), 403

    @app.errorhandler(404)
    def not_found(e):
        return render_template("error.html", code=404, message="Page Not Found"), 404

    return app


if __name__ == "__main__":
    app = create_app()
    app.run(debug=False, host='0.0.0.0', port=5000)
