import json
import re

from methodology.base import TakeoffContext, ProposedLineItem, ModuleResponse

SCOPE_CODE = "soldier_pile"
SCOPE_NAME = "Steel Soldier Pile Wall"

# Per Section 2.2 of David's methodology, soldier pile work expects these drawing types
EXPECTED_SHEETS = [
    "General Plan",
    "Structure Plan",
    "Retaining Wall Details",
    "Return Wall Details",
    "Lagging Details",
]


def opening_report(ctx: TakeoffContext) -> ModuleResponse:
    """
    Runs when the methodology takeoff page first loads.
    Surveys the project's Library + drawings and reports what's relevant
    for Soldier Pile work. Asks the estimator if anything is missing.
    """
    query = (
        "soldier pile structure plan pile data table retaining wall details "
        "return wall details lagging details"
    )
    retrieval = ctx.build_context_fn(ctx.project, query, ctx.processed_folder)

    found_sheets = []
    for entry in retrieval.get("index_map", {}).values():
        found_sheets.append(f"{entry['filename']} (page {entry['page']})")

    if not found_sheets:
        msg = (
            f"Soldier Pile takeoff for {ctx.project.name}. I don't see any drawings "
            f"loaded for this project yet. Per the methodology I'll need at minimum the "
            f"Structure Plan with the Pile Data Table. Can you upload the drawing set, "
            f"or point me to where it lives?"
        )
    else:
        sheets_summary = "\n".join(f"  • {s}" for s in found_sheets[:10])
        msg = (
            f"Soldier Pile takeoff for {ctx.project.name}. Here's what I see in the "
            f"project library:\n\n{sheets_summary}\n\n"
            f"For Soldier Pile work the methodology expects: Structure Plan (with Pile "
            f"Data Table), Retaining Wall Details, Return Wall Details, and Lagging "
            f"Details. If any of those aren't in the list above, let me know where to "
            f"find them. Otherwise, ready to map the pile inventory when you are."
        )

    return ModuleResponse(
        message=msg,
        sources=found_sheets,
        used_fallback=retrieval.get("used_fallback", False),
    )


def propose_step_1_inventory(ctx: TakeoffContext, user_message: str = "") -> ModuleResponse:
    """
    Step 1 — Map the 3D Structure (here: enumerate the piles).
    Asks Claude Vision to read the Pile Data Table and propose one line item per
    pile group (sections + counts + heights driven by lagging count).
    """
    query = (
        "Read the Pile Data Table on the Structure Plan. For each pile or group of "
        "identical piles, identify: pile number or station, steel section (HP14x102, "
        "HP14x89, W18x119, etc.), and the lagging count for each adjacent bay. "
        "Apply the lagging count rule from Section 2.4: painted pile height = the "
        "HIGHER of the two adjacent bay lagging counts (each lagging board = 1 ft of "
        "exposed height). End piles use the single bay count."
    )
    if user_message:
        query = f"{user_message}\n\nALSO: {query}"

    retrieval = ctx.build_context_fn(ctx.project, query, ctx.processed_folder)

    content = list(retrieval["content_blocks"])
    content.append({
        "type": "text",
        "text": (
            f"You are assisting with a structured Soldier Pile takeoff for the "
            f"project '{ctx.project.name}'. The images above are drawing pages.\n\n"
            f"Methodology context:\n{retrieval['text_context']}\n\n"
            f"TASK: {query}\n\n"
            f"Return a JSON array of line items. Each item has these fields:\n"
            f"  - element: a short label (e.g. 'HP14x102 Piles 1-12')\n"
            f"  - qty: number of identical piles in this group\n"
            f"  - height_ft: painted pile height in feet (from lagging count rule)\n"
            f"  - dwg_ref: which page / drawing you read this from\n"
            f"  - notes: any assumptions, scope flags, or field-verify items\n\n"
            f"Group consecutive identical piles together. Do NOT propose factors or "
            f"SF/LF values yet — that's Step 2. Just inventory and heights.\n\n"
            f"Return ONLY the JSON array, no preamble. If you cannot read the Pile "
            f"Data Table from the drawings provided, return an empty array [] and "
            f"explain in a separate JSON field 'unable_reason'."
        ),
    })

    message = ctx.anthropic_client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4096,
        messages=[{"role": "user", "content": content}],
    )
    raw_response = message.content[0].text

    proposed = []
    parse_note = ""
    try:
        json_match = re.search(r"\[\s*(?:\{.*?\}\s*,?\s*)*\]", raw_response, re.DOTALL)
        if json_match:
            parsed = json.loads(json_match.group(0))
            for idx, item in enumerate(parsed):
                proposed.append(ProposedLineItem(
                    step=1,
                    sort_order=idx,
                    element=item.get("element"),
                    qty=item.get("qty"),
                    height_ft=item.get("height_ft"),
                    dwg_ref=item.get("dwg_ref"),
                    notes=item.get("notes"),
                ))
        else:
            parse_note = "Couldn't extract a JSON array from Skippy's response."
    except (json.JSONDecodeError, Exception) as e:
        parse_note = f"JSON parsing failed: {e}"

    if proposed:
        msg = (
            f"I've read the drawings and propose {len(proposed)} pile group(s) "
            f"for Step 1. Review the inventory below — accept what looks right, "
            f"edit what doesn't, and let me know if anything is missing."
        )
    else:
        msg = (
            f"I couldn't extract a structured pile inventory from the drawings I "
            f"have access to. {parse_note}\n\nClaude's raw response was:\n\n"
            f"{raw_response[:1000]}"
        )

    return ModuleResponse(
        message=msg,
        proposed_items=proposed,
        sources=list(retrieval.get("index_map", {}).values()),
        used_fallback=retrieval.get("used_fallback", False),
    )


def propose_step_2_sizes_factors(ctx: TakeoffContext, user_message: str = "") -> ModuleResponse:
    """Step 2 — Assign Member Sizes and Surface Area Factors. NOT YET IMPLEMENTED."""
    return ModuleResponse(
        message="Step 2 (Sizes & SF/LF Factors) is not yet implemented in this build."
    )


def propose_step_3_adjustments(ctx: TakeoffContext, user_message: str = "") -> ModuleResponse:
    """Step 3 — Identify and Adjust for Non-Exposed Surfaces. NOT YET IMPLEMENTED."""
    return ModuleResponse(
        message="Step 3 (Non-Exposed Surface Adjustments) is not yet implemented in this build."
    )


def propose_step_4_calculate(ctx: TakeoffContext, user_message: str = "") -> ModuleResponse:
    """Step 4 — Calculate Surface Area. NOT YET IMPLEMENTED."""
    return ModuleResponse(
        message="Step 4 (Calculate) is not yet implemented in this build."
    )
