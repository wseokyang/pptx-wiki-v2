"""Structure-first PPTX to wiki conversion."""

from .models import BBox, DeckRecord, Element, SlideRecord, TableCell, TableData
from .extract import extract_pptx
from .config import AppConfig, load_config
from .configured import run_configured
from .pipeline import PipelineConfig, PipelineResult, run_pipeline

__all__ = [
    "BBox",
    "AppConfig",
    "DeckRecord",
    "Element",
    "PipelineConfig",
    "PipelineResult",
    "SlideRecord",
    "TableCell",
    "TableData",
    "extract_pptx",
    "load_config",
    "run_configured",
    "run_pipeline",
]

__version__ = "0.1.0"
