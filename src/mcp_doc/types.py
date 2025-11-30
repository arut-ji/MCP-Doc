"""Type definitions for the MCP Docx Processing Service"""

from typing import Literal, TypedDict, Union, Optional, List, Any


class RunInfo(TypedDict, total=False):
    """Type definition for run formatting information"""
    bold: Optional[bool]
    italic: Optional[bool]
    underline: Optional[bool]
    font_size: Optional[Any]  # Pt type from docx.shared
    font_name: Optional[str]
    color: Optional[Any]  # RGBColor type from docx.shared


class StyleInfo(TypedDict, total=False):
    """Type definition for paragraph style information"""
    style: Optional[Any]  # docx style object
    alignment: Optional[Any]  # WD_PARAGRAPH_ALIGNMENT enum
    runs: List[RunInfo]


class SearchResultParagraph(TypedDict):
    """Type definition for paragraph search result"""
    type: Literal["paragraph"]
    index: int
    text: str


class SearchResultTableCell(TypedDict):
    """Type definition for table cell search result"""
    type: Literal["table cell"]
    table_index: int
    row: int
    column: int
    text: str


SearchResult = Union[SearchResultParagraph, SearchResultTableCell]


class ReplaceResultParagraph(TypedDict):
    """Type definition for paragraph replace result"""
    type: Literal["paragraph"]
    index: int
    original: str
    replaced: str
    count: int


class ReplaceResultTableCell(TypedDict):
    """Type definition for table cell replace result"""
    type: Literal["table cell"]
    table_index: int
    row: int
    column: int
    original: str
    replaced: str
    count: int


ReplaceResult = Union[ReplaceResultParagraph, ReplaceResultTableCell]

