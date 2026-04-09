import anthropic

from models import db, Drawing, DrawingPage, Project, Company


def search_drawings(query, company_id, api_key):
    """Use Claude API to search through extracted drawing text."""
    # Gather all completed drawing text for the company
    pages = (
        db.session.query(DrawingPage, Drawing, Project)
        .join(Drawing, DrawingPage.drawing_id == Drawing.id)
        .join(Project, Drawing.project_id == Project.id)
        .filter(Project.company_id == company_id)
        .filter(Drawing.status == "completed")
        .filter(DrawingPage.extracted_text != "")
        .all()
    )

    if not pages:
        return {
            "answer": "No processed drawings found for your company. Please upload and process some drawings first.",
            "sources": [],
        }

    # Build context from drawing pages
    context_parts = []
    source_map = {}
    for page, drawing, project in pages:
        key = f"[{project.name} / {drawing.original_filename} / Page {page.page_number}]"
        context_parts.append(f"{key}\n{page.extracted_text}\n")
        source_map[key] = {
            "drawing_id": drawing.id,
            "project": project.name,
            "filename": drawing.original_filename,
            "page": page.page_number,
        }

    context = "\n---\n".join(context_parts)

    # Truncate context if too large (keep under ~150k chars for safety)
    if len(context) > 150000:
        context = context[:150000] + "\n...[truncated]"

    client = anthropic.Anthropic(api_key=api_key)

    message = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=2048,
        messages=[
            {
                "role": "user",
                "content": f"""You are an engineering drawing search assistant. Below is extracted OCR text from engineering drawings. Answer the user's question based on this content. Reference specific drawings and page numbers in your answer.

DRAWING CONTENT:
{context}

USER QUESTION: {query}

Provide a clear, specific answer referencing the relevant drawings. If the information is not found in the drawings, say so clearly.""",
            }
        ],
    )

    answer = message.content[0].text

    # Extract referenced sources from the answer
    sources = []
    for key, info in source_map.items():
        if key in answer or info["filename"] in answer or info["project"] in answer:
            sources.append(info)

    return {"answer": answer, "sources": sources[:10]}
