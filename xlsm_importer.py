"""Parser for the CCC Estimate Blank workbook (.xlsm / .xlsx)."""

import logging

import openpyxl

logger = logging.getLogger(__name__)

# Name of the worksheet that contains the estimate inputs.
SHEET_NAME = "Estimate"

# Cell addresses verified from the Estimate sheet's own formulas.
# None = not present on the sheet; field stays yellow on the review screen.
CELL_MAP: dict[str, str | None] = {
    # ── Project Parameters ─────────────────────────────────────────────────
    "deck_area_sf":           "B6",    # primary deck area input; referenced by =B6 across sheet
    "blast_level":            "C8",    # text like "SP-7", "SP-10"
    "abrasive_type":          "G9",    # text like "Steel grit G-25"
    "abrasive_lb_per_sf":     "E9",    # used in tons formula =E9*D8/2000

    # ── Materials (col F = VOL%, col H = MILS, col J = GAL) ───────────────
    # Gallons formula: =+D{row}/1604*H{row}/(F{row}+0.0001)*100
    "primer_vol_pct":                  "F11",
    "primer_mils":                     "H11",
    "primer_gal":                      "J11",
    "second_primer_vol_pct":           "F12",
    "second_primer_mils":              "H12",
    "second_primer_gal":               "J12",
    "stripe_prime_vol_pct":            "F13",
    "stripe_prime_mils":               "H13",
    "stripe_prime_gal":                "J13",
    "stripe_intermediate_vol_pct":     "F14",
    "stripe_intermediate_mils":        "H14",
    "stripe_intermediate_gal":         "J14",
    "intermediate_vol_pct":            "F15",
    "intermediate_mils":               "H15",
    "intermediate_gal":                "J15",
    "finish_vol_pct":                  "F16",
    "finish_mils":                     "H16",
    "finish_gal":                      "J16",

    # ── Labor — time-based (col D = HRS/DAY, col F = DAYS for rows 21-24) ─
    "mobilize_hrs_per_day":            "D21",
    "mobilize_days":                   "F21",
    "equip_setup_hrs_per_day":         "D23",
    "equip_setup_days":                "F23",
    "scaffold_hrs_per_day":            "D24",
    "scaffold_days":                   "F24",
    # rows 38-40: col F = HRS/DAY, col H = DAYS
    "traffic_control_hrs_per_day":     "F38",
    "traffic_control_days":            "H38",
    "inspection_touchup_hrs_per_day":  "F39",
    "inspection_touchup_days":         "H39",
    "osha_training_hrs_per_day":       "F40",
    "osha_training_days":              "H40",

    # ── Labor — production tasks (col D = SF/HR, col H = workers/nozzle) ──
    "containment_sf_per_hr":               "D25",
    "containment_workers_per_nozzle":      None,   # no workers/nozzle multiplier on sheet
    "masking_sf_per_hr":                   "D26",
    "masking_workers_per_nozzle":          "H26",
    "pressure_wash_sf_per_hr":             "D27",
    "pressure_wash_workers_per_nozzle":    "H27",
    "caulking_sf_per_hr":                  "D28",  # sheet uses LF/HR; stored as sf_per_hr by convention
    "caulking_workers_per_nozzle":         None,   # not on sheet for caulking
    "blast_sf_per_hr":                     "D29",
    "blast_workers_per_nozzle":            "H29",
    "primer_labor_sf_per_hr":              "D30",
    "primer_labor_workers_per_nozzle":     "H30",
    "second_primer_labor_sf_per_hr":       "D31",
    "second_primer_labor_workers_per_nozzle": "H31",
    "stripe_prime_labor_sf_per_hr":        "D32",
    "stripe_prime_labor_workers_per_nozzle": "H32",
    "stripe_intermediate_labor_sf_per_hr": "D33",
    "stripe_intermediate_labor_workers_per_nozzle": "H33",
    "intermediate_labor_sf_per_hr":        "D34",
    "intermediate_labor_workers_per_nozzle": "H34",
    "finish_labor_sf_per_hr":              "D35",
    "finish_labor_workers_per_nozzle":     "H35",

    # time-based tasks have no sf_per_hr / workers_per_nozzle on this sheet
    "mobilize_sf_per_hr":              None,
    "mobilize_workers_per_nozzle":     None,
    "equip_setup_sf_per_hr":           None,
    "equip_setup_workers_per_nozzle":  None,
    "scaffold_sf_per_hr":              None,
    "scaffold_workers_per_nozzle":     None,
    "traffic_control_sf_per_hr":       None,
    "traffic_control_workers_per_nozzle": None,
    "inspection_touchup_sf_per_hr":    None,
    "inspection_touchup_workers_per_nozzle": None,
    "osha_training_sf_per_hr":         None,
    "osha_training_workers_per_nozzle": None,

    # ── Shift Structure ────────────────────────────────────────────────────
    "shift_hours_per_day": "D79",   # sheet default 8
    "shift_days_total":    None,    # not on sheet; computed elsewhere
    "crew_size":           "B79",   # sheet default 10 (MEN)
    "shifts_per_day":      None,    # not on sheet; CCC template assumes single shift
}

# Fields that hold integer values (crew count, shifts per day).
_INT_FIELDS = {"crew_size", "shifts_per_day"}

# Fields that hold free-text strings rather than numbers.
_STR_FIELDS = {"blast_level", "abrasive_type"}

# Numeric fields that need 3 decimal places.
_THREE_DP_FIELDS = {"abrasive_lb_per_sf"}


def _read_cell(ws, addr: str):
    """Read one cell from the worksheet; return None for empty or error values."""
    try:
        val = ws[addr].value
        if val is None:
            return None
        if isinstance(val, str):
            val = val.strip()
            # openpyxl returns formula-error strings like '#REF!', '#VALUE!'
            if not val or val.startswith("#"):
                return None
        return val
    except Exception as exc:
        logger.warning("Could not read cell %s: %s", addr, exc)
        return None


def _coerce(field: str, raw):
    """Cast a raw cell value to the correct Python type for this field."""
    if raw is None:
        return None

    if field in _STR_FIELDS:
        return str(raw).strip() or None

    if field in _INT_FIELDS:
        try:
            return int(round(float(raw)))
        except (ValueError, TypeError):
            logger.warning("Non-integer value for %s: %r", field, raw)
            return None

    # Default: numeric
    scale = 3 if field in _THREE_DP_FIELDS else 2
    try:
        return round(float(raw), scale)
    except (ValueError, TypeError):
        logger.warning("Non-numeric value for %s: %r", field, raw)
        return None


def parse_estimate_workbook(file_stream) -> dict:
    """
    Parse a CCC Estimate Blank XLSM workbook from a file-like object.

    Returns a dict {field_name: value_or_None} for every field in CELL_MAP.
    Fields whose address is None in CELL_MAP always return None.
    Any cell that cannot be parsed returns None and logs a warning.
    Raises ValueError if the 'Estimate' sheet is not found.
    Raises openpyxl exceptions if the file cannot be opened at all.
    """
    result: dict = {field: None for field in CELL_MAP}

    wb = openpyxl.load_workbook(file_stream, read_only=True, data_only=True, keep_vba=False)

    if SHEET_NAME not in wb.sheetnames:
        wb.close()
        raise ValueError(
            f"This does not appear to be a CCC Estimate workbook — "
            f"sheet '{SHEET_NAME}' not found."
        )

    ws = wb[SHEET_NAME]

    for field, addr in CELL_MAP.items():
        if addr is None:
            continue
        raw = _read_cell(ws, addr)
        result[field] = _coerce(field, raw)

    logger.info(
        "Parsed %d/%d fields from sheet %r",
        sum(1 for v in result.values() if v is not None),
        len(result),
        SHEET_NAME,
    )

    wb.close()
    return result
