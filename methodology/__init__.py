from methodology import soldier_pile

_MODULES = {
    "soldier_pile": soldier_pile,
}


def get_module(scope_code: str):
    """Return the methodology module for a given scope_code, or None if not implemented."""
    return _MODULES.get(scope_code)


def available_scope_codes() -> list[str]:
    """Return list of scope_codes with implemented modules."""
    return list(_MODULES.keys())
