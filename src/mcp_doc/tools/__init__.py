"""MCP Tool Functions Registration"""

from mcp.server.fastmcp import FastMCP

# Import all tool functions
from .document import (
    create_document,
    open_document,
    save_document,
    save_as_document,
    create_document_copy,
    get_document_info,
)

from .content import (
    add_paragraph,
    add_heading,
    delete_paragraph,
    delete_text,
    search_text,
    search_and_replace,
    find_and_replace,
    replace_section,
    edit_section_by_keyword,
)

from .table import (
    add_table,
    add_table_row,
    delete_table_row,
    edit_table_cell,
    merge_table_cells,
    split_table,
)

from .layout import (
    add_page_break,
    set_page_margins,
)


def register_tools(mcp: FastMCP) -> None:
    """Register all tool functions with the MCP server"""
    # Document management tools
    mcp.tool()(create_document)
    mcp.tool()(open_document)
    mcp.tool()(save_document)
    mcp.tool()(save_as_document)
    mcp.tool()(create_document_copy)
    mcp.tool()(get_document_info)
    
    # Content tools
    mcp.tool()(add_paragraph)
    mcp.tool()(add_heading)
    mcp.tool()(delete_paragraph)
    mcp.tool()(delete_text)
    mcp.tool()(search_text)
    mcp.tool()(search_and_replace)
    mcp.tool()(find_and_replace)
    mcp.tool()(replace_section)
    mcp.tool()(edit_section_by_keyword)
    
    # Table tools
    mcp.tool()(add_table)
    mcp.tool()(add_table_row)
    mcp.tool()(delete_table_row)
    mcp.tool()(edit_table_cell)
    mcp.tool()(merge_table_cells)
    mcp.tool()(split_table)
    
    # Layout tools
    mcp.tool()(add_page_break)
    mcp.tool()(set_page_margins)
