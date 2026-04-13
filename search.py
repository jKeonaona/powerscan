import base64
import os
import re

import anthropic
from dotenv import load_dotenv

from models import db, Drawing, DrawingPage, Project

MAX_IMAGES_PER_REQUEST = 80


def search_drawings(query, project_id, api_key, processed_folder, doc_type=None):
    """Send all page images from a project to Claude Vision with the query."""
    load_dotenv()
    api_key = os.getenv("ANTHROPIC_API_KEY")
    print(f"[powerscan] search_drawings api_key: {api_key[:20] if api_key else '(empty)'}", flush=True)
    project = db.session.get(Project, project_id)
    if not project:
        return {"answer": "Project not found.", "sources": []}

    pages_q = (
        db.session.query(DrawingPage, Drawing)
        .join(Drawing, DrawingPage.drawing_id == Drawing.id)
        .filter(Drawing.project_id == project_id)
        .filter(Drawing.status == "ready")
    )
    if doc_type:
        pages_q = pages_q.filter(Drawing.doc_type == doc_type)
    pages = pages_q.order_by(Drawing.original_filename, DrawingPage.page_number).all()

    if not pages:
        scope = f" of type '{doc_type}'" if doc_type else ""
        return {
            "answer": f"No ready drawings{scope} were found in this project.",
            "sources": [],
        }

    if len(pages) > MAX_IMAGES_PER_REQUEST:
        return {
            "answer": (
                f"This project has {len(pages)} pages, exceeding the per-request limit "
                f"of {MAX_IMAGES_PER_REQUEST}. Narrow the scope or split the project."
            ),
            "sources": [],
        }

    content = []
    index_map = {}
    for i, (page, drawing) in enumerate(pages, start=1):
        abs_path = os.path.join(processed_folder, page.image_path)
        try:
            with open(abs_path, "rb") as f:
                data = base64.standard_b64encode(f.read()).decode("utf-8")
        except FileNotFoundError:
            continue

        label = f"INDEX {i}: {drawing.original_filename} — page {page.page_number}"
        index_map[i] = {
            "drawing_id": drawing.id,
            "filename": drawing.original_filename,
            "page": page.page_number,
        }
        content.append({"type": "text", "text": label})
        content.append({
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": "image/jpeg",
                "data": data,
            },
        })

    if not content:
        return {"answer": "No page images could be loaded from disk.", "sources": []}

    content.append({
        "type": "text",
        "text": (
            f"QUESTION: {query}\n\n"
            "You are an engineering drawing assistant. The images above are pages from "
            f"the project '{project.name}'. Each image is preceded by an 'INDEX N' label "
            "identifying its filename and page number. Study the actual drawing images "
            "(title blocks, notes, dimensions, schematics, labels) and answer the question.\n\n"
            "Your response MUST:\n"
            "1. Start with a direct answer.\n"
            "2. Cite the INDEX number(s) where you found the information, like [INDEX 3].\n"
            "3. If the information is not visible in any page, say so clearly."
        ),
    })

    client = anthropic.Anthropic(api_key=api_key)
    message = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=2048,
        messages=[{"role": "user", "content": content}],
    )
    answer = message.content[0].text

    sources = []
    seen = set()
    for match in re.findall(r"INDEX\s+(\d+)", answer):
        idx = int(match)
        if idx in index_map and idx not in seen:
            seen.add(idx)
            sources.append(index_map[idx])

    return {"answer": answer, "sources": sources}
