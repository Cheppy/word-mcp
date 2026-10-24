"""Measure and fix the layout of word processor documents over MCP."""

from .engine import DocumentEngine, EngineError, NoDocumentError
from .service import DocumentService

__all__ = ["DocumentEngine", "DocumentService", "EngineError", "NoDocumentError"]
__version__ = "0.3.0"
