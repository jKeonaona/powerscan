import base64
import os
import re
from collections import OrderedDict

import anthropic
from dotenv import load_dotenv

from models import db, Drawing, DrawingPage, Project

MAX_IMAGES_PER_REQUEST = 80


def _sample_pages(pages, limit):
    """Distribute `limit` slots across drawings so every drawing is represented.

    `pages` is a list of (DrawingPage, Drawing) tuples, pre-sorted by drawing
    filename and page number. Returns a new list of the same shape, sampled
    evenly within each drawing.
    """
    by_drawing = OrderedDict()
    for page, drawing in pages:
        by_drawing.setdefault(drawing.id, {"drawing": drawing, "pages": []})["pages"].append(page)

    drawings = list(by_drawing.values())
    n_drawings = len(drawings)

    if n_drawings >= limit:
        # Too many drawings to give each one a page — spread across drawings,
        # taking the middle page of each selected drawing.
        sampled = []
        for i in range(limit):
            idx = round(i * (n_drawings - 1) / max(1, limit - 1)) if limit > 1 else 0
            entry = drawings[idx]
            mid = entry["pages"][len(entry["pages"]) // 2]
            sampled.append((mid, entry["drawing"]))
        return sampled

    # Give each drawing at least one slot, then distribute the rest
    # proportionally to drawing size (biased toward longer documents).
    allocations = [1] * n_drawings
    remaining = limit - n_drawings
    if remaining > 0:
        extra_weights = [max(0, len(d["pages"]) - 1) for d in drawings]
        weight_sum = sum(extra_weights)
        if weight_sum > 0:
            for i in range(n_drawings):
                allocations[i] += round(remaining * extra_weights[i] / weight_sum)

    # Cap each allocation at the drawing's actual page count.
    for i, d in enumerate(drawings):
        allocations[i] = min(allocations[i], len(d["pages"]))

    # Trim/pad due to rounding so we land exactly on the limit.
    while sum(allocations) > limit:
        i = allocations.index(max(allocations))
        allocations[i] -= 1
    while sum(allocations) < limit:
        candidates = [i for i, d in enumerate(drawings) if allocations[i] < len(d["pages"])]
        if not candidates:
            break
        i = min(candidates, key=lambda k: allocations[k])
        allocations[i] += 1

    sampled = []
    for alloc, d in zip(allocations, drawings):
        pgs = d["pages"]
        if alloc <= 0:
            continue
        if alloc >= len(pgs):
            for p in pgs:
                sampled.append((p, d["drawing"]))
            continue
        if alloc == 1:
            indices = [len(pgs) // 2]
        else:
            indices = [round(j * (len(pgs) - 1) / (alloc - 1)) for j in range(alloc)]
        for idx in indices:
            sampled.append((pgs[idx], d["drawing"]))
    return sampled


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
        suggestion = (
            " Upload some drawings first, then try again."
            if not doc_type
            else " Try removing the document type filter or choose a different type."
        )
        return {
            "answer": f"No ready drawings{scope} were found in this project.{suggestion}",
            "sources": [],
        }

    total_pages = len(pages)
    sampled_notice = None
    if total_pages > MAX_IMAGES_PER_REQUEST:
        distinct_drawings = len({d.id for _, d in pages})
        pages = _sample_pages(pages, MAX_IMAGES_PER_REQUEST)
        filter_hint = (
            " For full coverage of a specific document type, use the type filter "
            "(Drawing, Contract, Specification, Bid Doc, Addendum, Other) to narrow the search."
            if not doc_type
            else " To cover every page, narrow your search further or split the project."
        )
        sampled_notice = (
            f"Note: this project has {total_pages} pages across {distinct_drawings} documents, "
            f"which is more than the {MAX_IMAGES_PER_REQUEST}-image limit per request. "
            f"I searched a representative sample of {len(pages)} pages distributed across every "
            f"document so each one is covered.{filter_hint}"
        )

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

    sampling_context = ""
    if sampled_notice:
        sampling_context = (
            "\n\nIMPORTANT: You are seeing a representative sample of pages drawn from "
            "every document in the project, not every page. If the answer depends on a "
            "page you cannot see, say so and recommend the user narrow the search (e.g. "
            "by document type) for full coverage."
        )

    content.append({
        "type": "text",
        "text": (
            f"QUESTION: {query}\n\n"
            "You are an engineering drawing assistant. The images above are pages from "
            f"the project '{project.name}'. Each image is preceded by an 'INDEX N' label "
            "identifying its filename and page number. Study the actual drawing images "
            "(title blocks, notes, dimensions, schematics, labels) and answer the question."
            f"{sampling_context}\n\n"
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

    if sampled_notice:
        answer = f"{sampled_notice}\n\n{answer}"

    return {"answer": answer, "sources": sources}
