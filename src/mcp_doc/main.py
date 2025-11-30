"""MCP Server setup and main entry point"""

import os
from contextlib import asynccontextmanager
from typing import AsyncIterator, Dict, Any

from mcp.server.fastmcp import FastMCP

from .config import CURRENT_DOC_FILE, logger
from .processor import processor
from .tools import register_tools


@asynccontextmanager
async def server_lifespan(server: FastMCP) -> AsyncIterator[Dict[str, Any]]:
    """Manage server lifecycle"""
    try:
        # Start server with clean state
        logger.info("DocxProcessor MCP server starting with clean state...")
        # Do not attempt to load any previous state
        yield {"processor": processor}
    finally:
        # Save state when server shuts down
        logger.info("DocxProcessor MCP server shutting down...")
        if processor.current_document and processor.current_file_path:
            processor.save_state()
        else:
            logger.info("No document open, not saving state")


# Create MCP server
mcp = FastMCP(
    name="DocxProcessor",
    instructions="Word document processing service, providing functions to create, edit, and query documents",
    lifespan=server_lifespan
)

# Register all tools
register_tools(mcp)


def main() -> None:
    """Main entry point for the MCP server."""
    # Always start with a clean state, don't try to load any previous document
    if os.path.exists(CURRENT_DOC_FILE):
        try:
            os.remove(CURRENT_DOC_FILE)
            logger.info("Removed existing state file for clean startup")
        except Exception as e:
            logger.error(f"Failed to remove existing state file: {e}")
    
    # Run MCP server
    mcp.run()


if __name__ == "__main__":
    main()

