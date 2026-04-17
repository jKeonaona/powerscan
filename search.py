import base64
import io
import os
import re

import anthropic
from dotenv import load_dotenv
from PIL import Image

from models import db, Drawing, DrawingPage, Project

MAX_IMAGES_PER_REQUEST = 20
MAX_IMAGE_WIDTH = 800
CLAUDE_MODEL = "claude-sonnet-4-20250514"


def _load_and_shrink(image_path):
    """Return base64-encoded JPEG bytes, resized to MAX_IMAGE_WIDTH if wider."""
    with Image.open(image_path) as img:
        if img.mode not in ("RGB", "L"):
            img = img.convert("RGB")
        if img.width > MAX_IMAGE_WIDTH:
            new_height = round(img.height * MAX_IMAGE_WIDTH / img.width)
            img = img.resize((MAX_IMAGE_WIDTH, new_height), Image.LANCZOS)
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=85, optimize=True)
    return base64.standard_b64encode(buf.getvalue()).decode("utf-8")


def _build_batch_content(batch_pages, processed_folder, start_index):
    """Build the (content, index_map) for one batch of pages, numbered from start_index."""
    content = []
    index_map = {}
    idx = start_index
    for page, drawing in batch_pages:
        abs_path = os.path.join(processed_folder, page.image_path)
        try:
            data = _load_and_shrink(abs_path)
        except FileNotFoundError:
            continue
        except Exception as e:
            print(f"[powerscan] skipping {abs_path}: {e}", flush=True)
            continue

        label = f"INDEX {idx}: {drawing.original_filename} — page {page.page_number}"
        index_map[idx] = {
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
        idx += 1
    return content, index_map


def _ask_batch(client, query, project, batch_pages, processed_folder, start_index,
               batch_num=None, total_batches=None, scope_context=None):
    """Send one batch of images to Claude Vision and return (answer_text, index_map)."""
    content, index_map = _build_batch_content(batch_pages, processed_folder, start_index)
    if not content:
        return "", {}

    if total_batches and total_batches > 1:
        batch_context = (
            f"\n\nThis is batch {batch_num} of {total_batches}. You are only seeing part of "
            f"the project. Answer based solely on the images in this batch; a later step will "
            f"combine your answer with the other batches. If the information is not present in "
            f"this batch, say exactly 'Not found in this batch.' so the combiner can ignore you."
        )
    else:
        batch_context = ""

    scope_prefix = f"Project work scope: {scope_context}\n\n" if scope_context else ""
    content.append({
        "type": "text",
        "text": (
            f"{scope_prefix}QUESTION: {query}\n\n"
            "You are an engineering drawing assistant. The images above are pages from "
            f"the project '{project.name}'. Each image is preceded by an 'INDEX N' label "
            "identifying its filename and page number. Study the actual drawing images "
            "(title blocks, notes, dimensions, schematics, labels) and answer the question."
            f"{batch_context}\n\n"
            "Your response MUST:\n"
            "1. Start with a direct answer (or 'Not found in this batch.').\n"
            "2. Cite the INDEX number(s) where you found the information, like [INDEX 3].\n"
            "3. If the information is not visible in any page, say so clearly."
        ),
    })

    message = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=2048,
        messages=[{"role": "user", "content": content}],
    )
    return message.content[0].text, index_map


def _synthesize(client, query, project, batch_answers, scope_context=None):
    """Combine per-batch answers into one coherent response via a text-only call."""
    sections = []
    for i, ans in enumerate(batch_answers, start=1):
        sections.append(f"--- BATCH {i} ANSWER ---\n{ans}")
    joined = "\n\n".join(sections)

    scope_prefix = f"Project work scope: {scope_context}\n\n" if scope_context else ""
    prompt = (
        f"{scope_prefix}You asked multiple batches of engineering drawings from the project "
        f"'{project.name}' the following question:\n\n"
        f"QUESTION: {query}\n\n"
        f"Each batch answered independently based only on the images it could see. "
        f"INDEX numbers are globally unique across all batches, so keep them as-is when "
        f"citing pages. Below are the per-batch answers.\n\n"
        f"{joined}\n\n"
        f"Combine these into ONE coherent answer. Rules:\n"
        f"1. Ignore any batch that said 'Not found in this batch.'\n"
        f"2. If batches agree, state the answer once.\n"
        f"3. If batches disagree or provide complementary details, reconcile or list them.\n"
        f"4. Preserve every useful [INDEX N] citation from the batch answers so the user "
        f"can trace the source.\n"
        f"5. If no batch found the information, say so clearly.\n"
        f"6. Do not mention 'batches' in your final answer — speak as if you read the whole "
        f"project directly."
    )

    message = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=2048,
        messages=[{"role": "user", "content": prompt}],
    )
    return message.content[0].text


def _extract_sources(answer, index_map):
    sources = []
    seen = set()
    for match in re.findall(r"INDEX\s+(\d+)", answer):
        idx = int(match)
        if idx in index_map and idx not in seen:
            seen.add(idx)
            sources.append(index_map[idx])
    return sources


def search_drawings(query, project_id, api_key, processed_folder, doc_type=None, scope_context=None):
    """Send all page images from a project to Claude Vision with the query.

    Projects with more than MAX_IMAGES_PER_REQUEST pages are split into batches,
    each asked independently, then the partial answers are synthesized into one
    coherent response.
    """
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

    client = anthropic.Anthropic(api_key=api_key)

    batches = [
        pages[i:i + MAX_IMAGES_PER_REQUEST]
        for i in range(0, len(pages), MAX_IMAGES_PER_REQUEST)
    ]
    total_batches = len(batches)
    print(f"[powerscan] search_drawings: {len(pages)} pages in {total_batches} batch(es)", flush=True)

    # Single-batch fast path: skip synthesis.
    if total_batches == 1:
        answer, index_map = _ask_batch(
            client, query, project, batches[0], processed_folder, start_index=1,
            scope_context=scope_context,
        )
        if not index_map:
            return {"answer": "No page images could be loaded from disk.", "sources": []}
        return {"answer": answer, "sources": _extract_sources(answer, index_map)}

    # Multi-batch: ask each batch, then synthesize.
    batch_answers = []
    full_index_map = {}
    next_index = 1
    for batch_num, batch in enumerate(batches, start=1):
        print(f"[powerscan] batch {batch_num}/{total_batches} ({len(batch)} pages)", flush=True)
        answer, index_map = _ask_batch(
            client, query, project, batch, processed_folder,
            start_index=next_index, batch_num=batch_num, total_batches=total_batches,
            scope_context=scope_context,
        )
        next_index += len(batch)
        if not index_map:
            continue
        full_index_map.update(index_map)
        batch_answers.append(answer)

    if not batch_answers:
        return {"answer": "No page images could be loaded from disk.", "sources": []}

    print(f"[powerscan] synthesizing {len(batch_answers)} batch answers", flush=True)
    final_answer = _synthesize(client, query, project, batch_answers, scope_context=scope_context)
    return {
        "answer": final_answer,
        "sources": _extract_sources(final_answer, full_index_map),
    }
