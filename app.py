import os
import uuid
import threading

from flask import (
    Flask, render_template, request, redirect, url_for, flash,
    send_from_directory, jsonify, abort,
)
from flask_login import (
    LoginManager, login_user, logout_user, login_required, current_user,
)

from config import Config
from models import (
    db, User, Company, Project, Drawing, DrawingPage,
    ROLE_SUPERADMIN, ROLE_ADMIN, ROLE_USER, ROLES,
)
from ocr import process_drawing, configure_tesseract
from search import search_drawings


def create_app():
    app = Flask(__name__)
    app.config.from_object(Config)

    os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
    os.makedirs(app.config["PROCESSED_FOLDER"], exist_ok=True)
    os.makedirs(os.path.join(app.instance_path), exist_ok=True)

    db.init_app(app)
    configure_tesseract(app.config["TESSERACT_CMD"])

    login_manager = LoginManager()
    login_manager.login_view = "login"
    login_manager.init_app(app)

    @login_manager.user_loader
    def load_user(user_id):
        return db.session.get(User, int(user_id))

    with app.app_context():
        db.create_all()
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
            return redirect(url_for("dashboard"))
        if request.method == "POST":
            username = request.form.get("username", "").strip()
            password = request.form.get("password", "")
            user = User.query.filter_by(username=username).first()
            if user and user.check_password(password):
                login_user(user)
                next_page = request.args.get("next")
                return redirect(next_page or url_for("dashboard"))
            flash("Invalid username or password.", "danger")
        return render_template("login.html")

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

    @app.route("/projects/<int:project_id>/drawings")
    @login_required
    def drawings(project_id):
        project = db.session.get(Project, project_id) or abort(404)
        if not current_user.is_superadmin and current_user.company_id != project.company_id:
            abort(403)
        return render_template("drawings.html", project=project)

    @app.route("/projects/<int:project_id>/upload", methods=["GET", "POST"])
    @login_required
    def upload_drawing(project_id):
        project = db.session.get(Project, project_id) or abort(404)
        if not current_user.is_superadmin and current_user.company_id != project.company_id:
            abort(403)

        if request.method == "POST":
            file = request.files.get("pdf_file")
            if not file or not file.filename.lower().endswith(".pdf"):
                flash("Please upload a valid PDF file.", "danger")
                return redirect(request.url)

            # Save with unique filename
            ext = os.path.splitext(file.filename)[1]
            unique_name = f"{uuid.uuid4().hex}{ext}"
            file.save(os.path.join(app.config["UPLOAD_FOLDER"], unique_name))

            drawing = Drawing(
                filename=unique_name,
                original_filename=file.filename,
                project_id=project.id,
                uploaded_by=current_user.id,
                status="pending",
            )
            db.session.add(drawing)
            db.session.commit()

            # Process OCR in background thread
            thread = threading.Thread(
                target=process_drawing, args=(app, drawing.id)
            )
            thread.daemon = True
            thread.start()

            flash(f"Drawing '{file.filename}' uploaded. OCR processing started.", "success")
            return redirect(url_for("drawings", project_id=project.id))

        return render_template("upload.html", project=project)

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
        return jsonify({"status": drawing.status})

    @app.route("/drawings/<int:drawing_id>/delete", methods=["POST"])
    @login_required
    @admin_required
    def delete_drawing(drawing_id):
        drawing = db.session.get(Drawing, drawing_id) or abort(404)
        project_id = drawing.project_id
        db.session.delete(drawing)
        db.session.commit()
        flash("Drawing deleted.", "success")
        return redirect(url_for("drawings", project_id=project_id))

    @app.route("/processed/<path:filename>")
    @login_required
    def serve_processed(filename):
        return send_from_directory(app.config["PROCESSED_FOLDER"], filename)

    # ── Search ──────────────────────────────────────────────────

    @app.route("/search", methods=["GET", "POST"])
    @login_required
    def search():
        results = None
        query = ""
        if request.method == "POST":
            query = request.form.get("query", "").strip()
            if not query:
                flash("Please enter a search query.", "danger")
            elif not app.config["ANTHROPIC_API_KEY"]:
                flash("Claude API key not configured. Set ANTHROPIC_API_KEY environment variable.", "danger")
            else:
                company_id = current_user.company_id
                if current_user.is_superadmin:
                    company_id = request.form.get("company_id", type=int) or Company.query.first().id if Company.query.first() else None
                if not company_id:
                    flash("No company selected.", "danger")
                else:
                    results = search_drawings(query, company_id, app.config["ANTHROPIC_API_KEY"])

        companies = Company.query.all() if current_user.is_superadmin else []
        return render_template("search.html", results=results, query=query, companies=companies)

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
    app.run(debug=True, port=5000)
