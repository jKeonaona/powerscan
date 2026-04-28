from dataclasses import dataclass, field
from typing import Optional, Callable


@dataclass
class TakeoffContext:
    """Bundle of dependencies passed from the route layer into module functions."""
    methodology_takeoff: object       # MethodologyTakeoff SQLAlchemy instance
    project: object                   # Project SQLAlchemy instance
    api_key: str
    processed_folder: str
    build_context_fn: Callable        # build_workspace_context from app.py
    anthropic_client: object          # anthropic.Anthropic instance


@dataclass
class ProposedLineItem:
    """Data dict for a line item the module proposes. Route layer persists these as MethodologyLineItem rows."""
    step: int
    sort_order: int = 0
    dwg_ref: Optional[str] = None
    description: Optional[str] = None
    element: Optional[str] = None
    qty: Optional[float] = None
    length_ft: Optional[float] = None
    height_ft: Optional[float] = None
    factor: Optional[float] = None
    notes: Optional[str] = None
    # total and sqft are computed by the route layer / DB layer, not the module


@dataclass
class ModuleResponse:
    """Standard return shape for module functions that produce conversation output."""
    message: str                                  # what Skippy says to the estimator
    proposed_items: list = field(default_factory=list)  # list[ProposedLineItem]
    sources: list = field(default_factory=list)         # citation list from Claude
    used_fallback: bool = False                         # did retrieval fall back to generic?
