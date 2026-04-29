import json
import re

from models import db, DrawingExtraction
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
        "Pile Data Table on Structure Plan with columns for pile number, station, "
        "steel section, pile length, and number of lagging. Soldier pile wall inventory."
    )
    if user_message:
        query = f"{user_message}\n\nALSO: {query}"

    retrieval = ctx.build_context_fn(ctx.project, query, ctx.processed_folder)

    # Find drawing_id from retrieval for cache lookup (first Drawing in index_map)
    index_map = retrieval.get("index_map", {})
    drawing_id = None
    if index_map:
        first_key = min(index_map.keys())
        drawing_id = index_map[first_key].get("drawing_id")

    # Cache hit — skip Vision entirely
    if drawing_id:
        cached = (
            DrawingExtraction.query
            .filter_by(drawing_id=drawing_id, scope_code=SCOPE_CODE)
            .order_by(DrawingExtraction.extraction_version.desc())
            .first()
        )
        if cached:
            proposed = []
            try:
                parsed = json.loads(cached.extracted_data_json)
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
            except (json.JSONDecodeError, Exception):
                pass
            cache_note = (
                f"Used cached extraction (version {cached.extraction_version} "
                f"from {cached.created_at.strftime('%Y-%m-%d %H:%M UTC')})."
            )
            if proposed:
                msg = (
                    f"I've read the drawings and propose {len(proposed)} pile group(s) "
                    f"for Step 1. Review the inventory below — accept what looks right, "
                    f"edit what doesn't, and let me know if anything is missing.\n\n"
                    f"{cache_note}"
                )
            else:
                msg = (
                    f"I couldn't extract a structured pile inventory from the cached data. "
                    f"{cache_note}"
                )
            return ModuleResponse(
                message=msg,
                proposed_items=proposed,
                sources=list(retrieval.get("index_map", {}).values()),
                used_fallback=retrieval.get("used_fallback", False),
            )

    # Cache miss — fire Vision
    content = list(retrieval["content_blocks"])
    content.append({
        "type": "text",
        "text": (
            f"You are reading a structural drawing set for a Soldier Pile takeoff "
            f"on the project '{ctx.project.name}'. The images above are drawing pages.\n\n"
            f"Methodology context:\n{retrieval['text_context']}\n\n"
            f"AUTHORITATIVE SOURCE: The Pile Data Table is the only acceptable source "
            f"for pile inventory. It is a tabular layout (rows = piles, columns = pile "
            f"attributes) located on a sheet titled 'Structure Plan' or similar. Common "
            f"columns include: Pile No., Station, Top of Wall Elevation, Top of Pile "
            f"Elevation, Steel Pile Type, Pile Length (LF), and Number of Lagging.\n\n"
            f"CRITICAL READING RULES:\n\n"
            f"1. READ EVERY ROW of the Pile Data Table. Do not summarize, skip, or "
            f"estimate. If a column value spans multiple cells as a merged cell with "
            f"one value, apply that value to every pile in the merged span.\n\n"
            f"2. DO NOT use bid item lists, plan notes, quantity sheets, or general "
            f"plan callouts as your source for pile sections. Pile sections come from "
            f"the 'Steel Pile Type' column in the Pile Data Table itself. Bid items "
            f"often only mention the most prominent sections and miss intermediate "
            f"sections that ARE in the table.\n\n"
            f"3. Read whatever section labels are actually written in the table. Do "
            f"not assume any particular sizes. Common Caltrans sections include "
            f"HP-shapes (HP10, HP12, HP14) and W-shapes (W12, W14, W18, W24), but the "
            f"specific sizes vary project by project.\n\n"
            f"4. The lagging column gives EXPOSED PILE HEIGHT in feet. Each lagging "
            f"board equals 1 ft of exposed wall height. Two tabulation styles exist:\n"
            f"   - PER PILE (one value in each pile's row): use that value directly as "
            f"the painted height.\n"
            f"   - PER BAY (one value between adjacent piles): painted height of pile "
            f"N = MAX(lagging in bay N-1 to N, lagging in bay N to N+1). End piles use "
            f"the single adjacent bay count.\n"
            f"   State explicitly in notes which style you found.\n\n"
            f"MULTI-SHEET CONTINUATION RULE: The Pile Data Table frequently spans multiple "
            f"drawing sheets. Look for sheets titled 'Structure Plan No. 1', 'Structure "
            f"Plan No. 2', 'Structure Plan No. 3' (and so on), or for MATCHLINE markers "
            f"on a sheet pointing to a continuation sheet, or for pile numbers in the table "
            f"continuing past where the first sheet ends. If ANY of these conditions are "
            f"true, you MUST read every sheet that contains Pile Data Table rows. "
            f"Concatenate the rows from all sheets into a single combined pile inventory. "
            f"Do NOT stop after the first sheet, even if its table looks self-contained. "
            f"Each sheet's pile data table is part of ONE wall inventory.\n\n"
            f"GROUPING RULE: Group consecutive piles into one line item ONLY when ALL "
            f"three of these are identical: steel section, pile length, and lagging "
            f"count. If any one of those differs from one pile to the next, start a "
            f"new line item. Err toward more line items, not fewer.\n\n"
            f"Return ONLY a JSON array. Each item has these fields:\n"
            f"  - element: short label (e.g. 'HP12x84 Piles 1-9, lagging 29')\n"
            f"  - qty: count of piles in this group\n"
            f"  - height_ft: painted pile height in feet (from lagging count rule)\n"
            f"  - length_ft: pile length in feet (from the Pile Length column)\n"
            f"  - dwg_ref: page reference (e.g. 'p.68')\n"
            f"  - notes: lagging tabulation style + any flags or assumptions\n\n"
            f"Do NOT propose factors or SF/LF values. That's Step 2.\n\n"
            f"If you cannot read the Pile Data Table, return an empty array [] and "
            f"add a separate top-level field 'unable_reason' explaining specifically "
            f"what was missing or unreadable."
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
    extracted_json_str = None
    try:
        json_match = re.search(r"\[\s*(?:\{.*?\}\s*,?\s*)*\]", raw_response, re.DOTALL)
        if json_match:
            extracted_json_str = json_match.group(0)
            parsed = json.loads(extracted_json_str)
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

    # Persist to cache on successful parse
    if drawing_id and extracted_json_str is not None:
        try:
            record = DrawingExtraction(
                drawing_id=drawing_id,
                scope_code=SCOPE_CODE,
                extraction_version=1,
                extracted_data_json=extracted_json_str,
                raw_vision_response=raw_response,
                manual_rerun=False,
                triggered_by_user_id=None,
                page_count_processed=len(retrieval["content_blocks"]),
                estimated_token_cost=None,
            )
            db.session.add(record)
            db.session.commit()
        except Exception:
            db.session.rollback()

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
