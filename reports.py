"""Document Intelligence report generator.

Takes a project's page images, runs them through Claude in batches using a
template prompt, synthesizes the batch responses into one structured payload,
and renders a branded .docx file via python-docx.
"""
import io
import json
import os
import re
from datetime import datetime, timezone

import anthropic
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from models import db, Drawing, DrawingPage, Project
from search import (
    CLAUDE_MODEL,
    MAX_IMAGES_PER_REQUEST,
    _build_batch_content,
)

# ── CCC brand constants (swap for real logo path later if desired) ──
BRAND_NAME = "CCC"
BRAND_TAGLINE = "Construction Intelligence Report"
BRAND_COLOR = RGBColor(0x1F, 0x4E, 0x79)   # deep blue
BRAND_ACCENT = RGBColor(0x8C, 0x8C, 0x8C)  # mid gray

SCHEMA_DESCRIPTION = """{
  "title": "string — report title",
  "summary": "string — 2-4 sentence executive summary",
  "sections": [
    // each entry is ONE of these section shapes:
    {"type": "heading", "text": "string", "level": 1-3},
    {"type": "paragraph", "text": "string"},
    {"type": "bullets", "items": ["string", ...]},
    {"type": "table", "headers": ["col1", "col2", ...], "rows": [["a","b"], ...]}
  ]
}"""

REPORT_TEMPLATES = {
    "submittal_list": {
        "name": "Submittal List",
        "description": "Every submittal required by the contract documents.",
        "prompt": (
            "Identify every submittal required by the project documents. For each submittal, "
            "extract: item name, specification section (if identifiable), submittal type "
            "(shop drawing / product data / sample / certificate / O&M manual / other), "
            "required timing or deadline if stated, and any special notes. "
            "Return results as a single table with columns: Item, Spec Section, Type, Timing, Notes. "
            "Start with a brief summary paragraph giving the total count and any submittals flagged "
            "as high-priority or potentially overdue."
        ),
    },
    "bid_summary": {
        "name": "Bid Summary",
        "description": "Contract overview, key dates, insurance, bonding, special provisions.",
        "prompt": (
            "Produce a bid summary covering: (1) contract/project overview, (2) key dates "
            "(bid due, award, NTP, substantial completion, final completion), (3) insurance "
            "requirements (types, limits, additional insureds), (4) bonding requirements "
            "(bid bond, performance bond, payment bond, percentages), (5) special provisions "
            "or unusual clauses the bidder must be aware of, (6) liquidated damages if any. "
            "Use headings and tables for dates, insurance, and bonding. Surface anything "
            "non-standard or unusually onerous in a clearly labeled section."
        ),
    },
    "compliance_checklist": {
        "name": "Compliance Checklist",
        "description": "Lead, environmental, and safety compliance requirements.",
        "prompt": (
            "Build a compliance checklist covering (1) lead compliance (RRP, EPA lead-safe, "
            "abatement requirements), (2) environmental requirements (hazardous materials, "
            "disposal, SWPPP, air quality, asbestos), (3) safety requirements "
            "(OSHA, site-specific safety plans, PPE, training). For each item, capture: "
            "requirement, source reference (spec section or page), and compliance action "
            "the contractor must take. Present as three tables — one per category — with "
            "columns: Requirement, Source, Action."
        ),
    },
    "risk_assessment": {
        "name": "Risk Assessment",
        "description": "Problematic clauses, non-standard terms, cost-impact risks.",
        "prompt": (
            "Produce a risk assessment flagging: (1) problematic clauses (unlimited "
            "indemnification, broad consequential damages, unusual warranty terms), "
            "(2) non-standard contract terms that deviate from AIA/ConsensusDocs norms, "
            "(3) potential cost impacts (unclear scope, coordination risks, schedule pressure, "
            "escalation exposure). For each risk, provide: description, severity "
            "(High / Medium / Low), source reference, and recommended mitigation. "
            "Present as a table with columns: Risk, Severity, Source, Mitigation. "
            "Start with an executive summary listing the top three risks."
        ),
    },
}


# ───────────────────────── Claude calls ─────────────────────────

def _parse_json_response(text):
    """Strip fences, extract first JSON object, parse. Fall back to paragraph."""
    text = (text or "").strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```\s*$", "", text)
    match = re.search(r"\{.*\}", text, re.DOTALL)
    if match:
        text = match.group(0)
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        return {
            "title": "Report",
            "summary": "",
            "sections": [{"type": "paragraph", "text": text}],
        }


def _ask_report_batch(client, project, template_prompt, batch_pages, processed_folder,
                     start_index, batch_num, total_batches):
    content, index_map = _build_batch_content(batch_pages, processed_folder, start_index)
    if not content:
        return None, {}

    batch_note = ""
    if total_batches > 1:
        batch_note = (
            f"\n\nThis is batch {batch_num} of {total_batches} from project "
            f"'{project.name}'. Report ONLY on what is visible in the images above. "
            f"A later step will merge your output with the other batches, so omit any "
            f"item you cannot confirm from this batch."
        )

    instructions = (
        f"{template_prompt}{batch_note}\n\n"
        f"Return a single JSON object matching this schema exactly:\n"
        f"{SCHEMA_DESCRIPTION}\n\n"
        f"Return ONLY the JSON. No markdown fences, no commentary, no leading or trailing text."
    )
    content.append({"type": "text", "text": instructions})

    message = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=4096,
        messages=[{"role": "user", "content": content}],
    )
    return _parse_json_response(message.content[0].text), index_map


def _synthesize_reports(client, project, template_name, template_prompt, batch_reports):
    """Combine multiple batch JSON reports into a single unified report JSON."""
    combined = json.dumps(batch_reports, indent=2)
    prompt = (
        f"You previously asked multiple batches of project documents (project: "
        f"'{project.name}') for a '{template_name}' report. Each batch returned its own "
        f"JSON report in this schema:\n{SCHEMA_DESCRIPTION}\n\n"
        f"TEMPLATE INTENT:\n{template_prompt}\n\n"
        f"PER-BATCH RESULTS:\n{combined}\n\n"
        f"Merge all batches into ONE final JSON report using the SAME schema. Rules:\n"
        f"1. Deduplicate items that appear in multiple batches.\n"
        f"2. Combine table rows with matching identifiers into a single table per category.\n"
        f"3. Reconcile any contradictions, noting the discrepancy if it cannot be resolved.\n"
        f"4. Preserve every unique finding.\n"
        f"5. Write a unified executive summary covering the whole project, not per-batch.\n"
        f"6. Do NOT mention 'batches' in the final output.\n\n"
        f"Return ONLY the final JSON. No markdown fences, no commentary."
    )
    message = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=8192,
        messages=[{"role": "user", "content": prompt}],
    )
    return _parse_json_response(message.content[0].text)


# ───────────────────────── DOCX rendering ─────────────────────────

def _set_cell_bg(cell, hex_color):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tc_pr.append(shd)


def _add_cover_page(doc, template_name, project_name, company_name):
    # Brand wordmark
    mark = doc.add_paragraph()
    mark.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = mark.add_run(BRAND_NAME)
    run.font.size = Pt(48)
    run.font.bold = True
    run.font.color.rgb = BRAND_COLOR

    tag = doc.add_paragraph()
    tag.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tag_run = tag.add_run(BRAND_TAGLINE.upper())
    tag_run.font.size = Pt(11)
    tag_run.font.color.rgb = BRAND_ACCENT
    tag_run.font.bold = True

    for _ in range(6):
        doc.add_paragraph()

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    t_run = title.add_run(template_name)
    t_run.font.size = Pt(32)
    t_run.font.bold = True
    t_run.font.color.rgb = BRAND_COLOR

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    s_run = sub.add_run(project_name)
    s_run.font.size = Pt(18)
    s_run.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

    if company_name:
        co = doc.add_paragraph()
        co.alignment = WD_ALIGN_PARAGRAPH.CENTER
        c_run = co.add_run(company_name)
        c_run.font.size = Pt(13)
        c_run.font.color.rgb = BRAND_ACCENT

    for _ in range(4):
        doc.add_paragraph()

    date_p = doc.add_paragraph()
    date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    d_run = date_p.add_run(datetime.now(timezone.utc).strftime("%B %d, %Y"))
    d_run.font.size = Pt(12)
    d_run.font.color.rgb = BRAND_ACCENT


def _style_headings(doc):
    for level, size in ((1, 18), (2, 14), (3, 12)):
        try:
            style = doc.styles[f"Heading {level}"]
            style.font.color.rgb = BRAND_COLOR
            style.font.size = Pt(size)
            style.font.bold = True
        except KeyError:
            pass


def _render_section(doc, section):
    stype = (section.get("type") or "paragraph").lower()

    if stype == "heading":
        level = max(1, min(3, int(section.get("level", 2) or 2)))
        doc.add_heading(str(section.get("text", "")), level=level)

    elif stype == "paragraph":
        doc.add_paragraph(str(section.get("text", "")))

    elif stype == "bullets":
        for item in section.get("items", []) or []:
            doc.add_paragraph(str(item), style="List Bullet")

    elif stype == "table":
        headers = section.get("headers", []) or []
        rows = section.get("rows", []) or []
        if not headers:
            return
        table = doc.add_table(rows=1 + len(rows), cols=len(headers))
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        try:
            table.style = "Light Grid Accent 1"
        except KeyError:
            pass

        hdr_cells = table.rows[0].cells
        for i, h in enumerate(headers):
            hdr_cells[i].text = str(h)
            _set_cell_bg(hdr_cells[i], "1F4E79")
            for para in hdr_cells[i].paragraphs:
                for run in para.runs:
                    run.bold = True
                    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

        for r_idx, row in enumerate(rows, start=1):
            cells = table.rows[r_idx].cells
            for c_idx in range(len(headers)):
                val = row[c_idx] if c_idx < len(row) else ""
                if isinstance(val, (list, dict)):
                    val = json.dumps(val)
                cells[c_idx].text = str(val) if val is not None else ""
        doc.add_paragraph()

    else:
        # Unknown type — render as paragraph so nothing is lost.
        text = section.get("text") or json.dumps(section)
        doc.add_paragraph(str(text))


def _render_report_docx(report, template_name, project):
    doc = Document()

    # Margins
    for section in doc.sections:
        section.top_margin = Inches(0.9)
        section.bottom_margin = Inches(0.9)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    _style_headings(doc)
    _add_cover_page(doc, template_name, project.name, project.company.name)
    doc.add_page_break()

    # Body header
    title = report.get("title") or template_name
    doc.add_heading(str(title), level=1)

    summary = (report.get("summary") or "").strip()
    if summary:
        doc.add_heading("Executive Summary", level=2)
        doc.add_paragraph(summary)
        doc.add_paragraph()

    sections = report.get("sections") or []
    if not sections:
        doc.add_paragraph(
            "No structured content was returned for this report.",
        )
    else:
        for section in sections:
            if isinstance(section, dict):
                _render_section(doc, section)

    # Footer / generated note
    doc.add_paragraph()
    footer = doc.add_paragraph()
    f_run = footer.add_run(
        f"Generated by {BRAND_NAME} · {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M UTC')}"
    )
    f_run.italic = True
    f_run.font.size = Pt(9)
    f_run.font.color.rgb = BRAND_ACCENT

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ───────────────────────── Public entry point ─────────────────────────

def generate_report(project_id, template_id, custom_prompt, api_key, processed_folder):
    """Run the full pipeline and return (docx_bytes, filename, template_name, summary)."""
    project = db.session.get(Project, project_id)
    if not project:
        raise ValueError("Project not found.")

    if template_id == "custom":
        template_name = "Custom Report"
        prompt = (custom_prompt or "").strip()
        if not prompt:
            raise ValueError("Custom report requires a prompt.")
    else:
        tpl = REPORT_TEMPLATES.get(template_id)
        if not tpl:
            raise ValueError(f"Unknown template: {template_id}")
        template_name = tpl["name"]
        prompt = tpl["prompt"]

    pages = (
        db.session.query(DrawingPage, Drawing)
        .join(Drawing, DrawingPage.drawing_id == Drawing.id)
        .filter(Drawing.project_id == project_id)
        .filter(Drawing.status == "ready")
        .order_by(Drawing.original_filename, DrawingPage.page_number)
        .all()
    )
    if not pages:
        raise ValueError("No ready drawings in this project. Upload and wait for conversion first.")

    client = anthropic.Anthropic(api_key=api_key)
    batches = [
        pages[i:i + MAX_IMAGES_PER_REQUEST]
        for i in range(0, len(pages), MAX_IMAGES_PER_REQUEST)
    ]
    total_batches = len(batches)
    print(f"[powerscan] generate_report {template_id}: {len(pages)} pages in {total_batches} batch(es)", flush=True)

    batch_reports = []
    next_index = 1
    for batch_num, batch in enumerate(batches, start=1):
        print(f"[powerscan] report batch {batch_num}/{total_batches}", flush=True)
        result, _ = _ask_report_batch(
            client, project, prompt, batch, processed_folder,
            start_index=next_index, batch_num=batch_num, total_batches=total_batches,
        )
        next_index += len(batch)
        if result:
            batch_reports.append(result)

    if not batch_reports:
        raise ValueError("No content returned from Claude for any batch.")

    if total_batches == 1:
        final_report = batch_reports[0]
    else:
        print(f"[powerscan] synthesizing {len(batch_reports)} batch reports", flush=True)
        final_report = _synthesize_reports(client, project, template_name, prompt, batch_reports)

    docx_bytes = _render_report_docx(final_report, template_name, project)

    safe_project = re.sub(r"[^A-Za-z0-9_-]+", "_", project.name).strip("_") or "project"
    safe_template = re.sub(r"[^A-Za-z0-9_-]+", "_", template_name).strip("_") or "report"
    filename = f"{safe_project}-{safe_template}-{datetime.now(timezone.utc).strftime('%Y%m%d-%H%M')}.docx"

    summary = (final_report.get("summary") or "").strip()
    return docx_bytes, filename, template_name, summary
