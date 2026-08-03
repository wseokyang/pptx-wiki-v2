"""Structure-first PPTX parsing, semantic reorganisation, and Wiki publishing."""

from .models import BBox, DeckRecord, Element, SlideRecord, TableCell, TableData
from .extract import extract_pptx
from .config import AppConfig, load_config
from .configured import run_configured
from .pipeline import PipelineConfig, PipelineResult, run_pipeline
from .semantic import (
    SemanticConfig,
    SemanticDocument,
    SemanticExport,
    build_semantic_output,
    load_semantic_documents,
)
from .wiki_publish import WikiExport, publish_wiki

__all__ = [
    "BBox",
    "AppConfig",
    "DeckRecord",
    "Element",
    "PipelineConfig",
    "PipelineResult",
    "SemanticConfig",
    "SemanticDocument",
    "SemanticExport",
    "SlideRecord",
    "TableCell",
    "TableData",
    "WikiExport",
    "build_semantic_output",
    "extract_pptx",
    "load_config",
    "load_semantic_documents",
    "publish_wiki",
    "run_configured",
    "run_pipeline",
]

__version__ = "0.2.0"
