"""Structure-first PPTX parsing, semantic reorganisation, and Wiki publishing."""

from .models import BBox, DeckRecord, Element, SlideRecord, TableCell, TableData
from .extract import extract_pptx
from .config import AppConfig, load_config
from .configured import run_configured, run_configured_collection
from .collection import (
    CollectionConfig,
    CollectionResult,
    CollectionSourceResult,
    InputOccurrence,
    discover_pptx_inputs,
    run_collection,
)
from .integration import (
    IntegratedExport,
    IntegrationConfig,
    build_integrated_artifact,
    validate_integrated_artifact,
)
from .pipeline import PipelineConfig, PipelineResult, run_pipeline
from .semantic import (
    SemanticConfig,
    SemanticDocument,
    SemanticExport,
    build_semantic_output,
    load_semantic_documents,
)
from .wiki_publish import WikiExport, publish_wiki
from .quartz_publish import QuartzExport, publish_quartz
from .source_semantic import (
    SourceIdentity,
    SourceSemanticExport,
    build_source_semantic,
    extract_pr_numbers,
    load_source_semantic,
)

__all__ = [
    "BBox",
    "AppConfig",
    "CollectionConfig",
    "CollectionResult",
    "CollectionSourceResult",
    "DeckRecord",
    "Element",
    "InputOccurrence",
    "IntegratedExport",
    "IntegrationConfig",
    "PipelineConfig",
    "PipelineResult",
    "QuartzExport",
    "SemanticConfig",
    "SemanticDocument",
    "SemanticExport",
    "SlideRecord",
    "SourceIdentity",
    "SourceSemanticExport",
    "TableCell",
    "TableData",
    "WikiExport",
    "build_semantic_output",
    "build_integrated_artifact",
    "build_source_semantic",
    "discover_pptx_inputs",
    "extract_pptx",
    "load_config",
    "load_semantic_documents",
    "load_source_semantic",
    "extract_pr_numbers",
    "publish_wiki",
    "publish_quartz",
    "run_configured",
    "run_configured_collection",
    "run_collection",
    "run_pipeline",
    "validate_integrated_artifact",
]

__version__ = "0.3.0"
