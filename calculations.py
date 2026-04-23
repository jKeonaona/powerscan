"""Takeoff calculation engine — quantity-only, no dollar conversion."""

_EPS = 0.0001  # divide-by-zero guard, matches Excel pattern

_COATING_CONSTANT = 1604  # industry-standard mils-to-gallons constant

_SUPERVISION_RATE = 0.12

_COATING_LAYERS = [
    "primer",
    "second_primer",
    "stripe_prime",
    "stripe_intermediate",
    "intermediate",
    "finish",
]

# Production tasks: driven by deck_area_sf / sf_per_hr * workers_per_nozzle.
# NOTE: caulking is technically LF-based, not SF; treated as SF here until
# a separate LF input is added in a future build.
_PRODUCTION_TASKS = [
    "pressure_wash",
    "caulking",
    "blast",
    "containment",
    "masking",
    "primer_labor",
    "second_primer_labor",
    "stripe_prime_labor",
    "stripe_intermediate_labor",
    "intermediate_labor",
    "finish_labor",
]

# Time-based tasks: driven by hrs_per_day * days.
_TIME_TASKS = [
    "mobilize",
    "equip_setup",
    "scaffold",
    "traffic_control",
    "inspection_touchup",
    "osha_training",
]

# Friendly labels used when serialising for the template.
_TASK_LABELS = {
    "mobilize":                    "Mobilize",
    "equip_setup":                 "Equip Setup",
    "scaffold":                    "Scaffold",
    "containment":                 "Containment",
    "masking":                     "Masking",
    "pressure_wash":               "Pressure Wash",
    "caulking":                    "Caulking",
    "blast":                       "Blast",
    "primer_labor":                "Primer",
    "second_primer_labor":         "2nd Primer",
    "stripe_prime_labor":          "Stripe Prime",
    "stripe_intermediate_labor":   "Stripe Intermediate",
    "intermediate_labor":          "Intermediate",
    "finish_labor":                "Finish",
    "traffic_control":             "Traffic Control",
    "inspection_touchup":          "Inspection / Touchup",
    "osha_training":               "OSHA Training",
}

# Display order for labor-hours table (time-based first, then production).
_LABOR_DISPLAY_ORDER = _TIME_TASKS + _PRODUCTION_TASKS


def _f(value):
    """Convert a SQLAlchemy Numeric/None to float, or None if None."""
    if value is None:
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def calculate_painting_quantities(takeoff):
    """
    Compute painting quantities from a Takeoff record's input fields.
    Pure function — never mutates the takeoff.
    Keys are always present in the returned dict; values may be None
    (meaning "not enough inputs to compute this quantity").
    """
    area = _f(takeoff.deck_area_sf)

    # ── Materials: gallons per coating layer ──────────────────────────────────
    materials = {}
    for layer in _COATING_LAYERS:
        mils = _f(getattr(takeoff, f"{layer}_mils"))
        vol_pct = _f(getattr(takeoff, f"{layer}_vol_pct"))
        if area is None or mils is None or mils == 0 or vol_pct is None or vol_pct < _EPS:
            materials[f"{layer}_gal"] = None
        else:
            materials[f"{layer}_gal"] = (area * mils / _COATING_CONSTANT) * (100.0 / vol_pct)

    # ── Abrasive tonnage ──────────────────────────────────────────────────────
    lb_per_sf = _f(takeoff.abrasive_lb_per_sf)
    if area is None or lb_per_sf is None:
        abrasive_tons = None
    else:
        abrasive_tons = (lb_per_sf * area) / 2000.0

    # ── Labor hours per production task ───────────────────────────────────────
    labor_hours = {}
    for task in _PRODUCTION_TASKS:
        sf_per_hr = _f(getattr(takeoff, f"{task}_sf_per_hr"))
        workers = _f(getattr(takeoff, f"{task}_workers_per_nozzle"))
        if area is None or sf_per_hr is None or sf_per_hr < _EPS or workers is None:
            labor_hours[task] = None
        else:
            labor_hours[task] = (area / sf_per_hr) * workers

    # ── Labor hours per time-based task ──────────────────────────────────────
    for task in _TIME_TASKS:
        hrs_day = _f(getattr(takeoff, f"{task}_hrs_per_day"))
        days = _f(getattr(takeoff, f"{task}_days"))
        if hrs_day is None or days is None:
            labor_hours[task] = None
        else:
            labor_hours[task] = hrs_day * days

    # ── Labor totals ──────────────────────────────────────────────────────────
    known_hours = [h for h in labor_hours.values() if h is not None]
    sum_task_hours = sum(known_hours) if known_hours else None
    if sum_task_hours is not None:
        supervision_hours = sum_task_hours * _SUPERVISION_RATE
        total_labor_hours = sum_task_hours + supervision_hours
    else:
        supervision_hours = None
        total_labor_hours = None

    # ── Schedule: work days ───────────────────────────────────────────────────
    crew = _f(takeoff.crew_size)
    shift_hrs = _f(takeoff.shift_hours_per_day)
    if crew is None or shift_hrs is None or (crew * shift_hrs) < _EPS or total_labor_hours is None:
        work_days = None
    else:
        work_days = total_labor_hours / (crew * shift_hrs)

    return {
        "materials": materials,
        "abrasive": {"tons": abrasive_tons},
        "labor_hours_by_task": labor_hours,
        "labor_totals": {
            "sum_task_hours": sum_task_hours,
            "supervision_hours": supervision_hours,
            "total_labor_hours": total_labor_hours,
        },
        "schedule": {"work_days": work_days},
        # Ordered task list for the template table
        "_task_display_order": _LABOR_DISPLAY_ORDER,
        "_task_labels": _TASK_LABELS,
        "_coating_layers": _COATING_LAYERS,
    }


def has_any_inputs(takeoff):
    """Return True if at least one calculation-relevant input field is set."""
    check_fields = (
        ["deck_area_sf", "abrasive_lb_per_sf", "crew_size", "shift_hours_per_day"]
        + [f"{layer}_mils" for layer in _COATING_LAYERS]
        + [f"{layer}_vol_pct" for layer in _COATING_LAYERS]
        + [f"{task}_sf_per_hr" for task in _PRODUCTION_TASKS]
        + [f"{task}_hrs_per_day" for task in _TIME_TASKS]
    )
    return any(getattr(takeoff, f, None) is not None for f in check_fields)
