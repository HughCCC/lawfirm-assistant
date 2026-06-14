"""Pydantic models for API request/response and JSON validation."""

from typing import Literal, Optional, Union
from pydantic import BaseModel, Field


# ─── API Request / Response ───────────────────────────────────────────

class GenerateRequest(BaseModel):
    """Request body for POST /api/generate."""
    question: str = Field(
        ...,
        min_length=20,
        description="User's legal question / scenario description",
    )
    provider: Literal["openai", "anthropic", "deepseek", "custom"] = Field(
        ...,
        description="LLM provider",
    )
    api_key: str = Field(..., min_length=10, description="API key for the provider")
    model: Optional[str] = Field(None, description="Model name override")
    base_url: Optional[str] = Field(None, description="Base URL for custom provider")


class GenerateResponse(BaseModel):
    """Response body for POST /api/generate."""
    success: bool
    title: str = ""
    docx_filename: str = ""
    docx_url: str = ""
    error: str = ""
    retries_used: int = 0


# ─── Report JSON Validation ───────────────────────────────────────────

class TableContent(BaseModel):
    """Content schema for table-type sections."""
    caption: str = Field(..., description="Table caption: 表N：标题")
    headers: list[str] = Field(..., min_length=1, description="Column headers")
    rows: list[list[str]] = Field(..., min_length=0, description="Data rows")
    note: Optional[str] = Field(None, description="Table footnote / note text")
    footnote_on_col: Optional[str] = Field(
        None, description="Column name that gets superscript footnote markers"
    )
    footnote_urls: Optional[list[str]] = Field(
        None, description="URLs for footnotes, one per row"
    )


class Section(BaseModel):
    """A single section in the report JSON."""
    type: Literal["h1", "h2", "h3", "h4", "h5", "body", "table", "source_item"]
    content: Union[str, TableContent]


class Report(BaseModel):
    """Full report JSON structure."""
    title: str = Field(..., min_length=1, description="Report title")
    sections: list[Section] = Field(..., min_length=1, description="Report sections")


class ValidationError(BaseModel):
    """Represents a single validation error."""
    field: str
    message: str
    section_index: Optional[int] = None


class SearchResult(BaseModel):
    """A single web search result with extracted case numbers."""
    title: str
    url: str
    snippet: str
    case_numbers: list[str] = Field(default_factory=list)
