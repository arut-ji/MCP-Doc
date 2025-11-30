"""Configuration, constants, and logging setup"""

import os
import tempfile
import logging

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler(os.path.join(tempfile.gettempdir(), "docx_mcp_server.log")),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger("DocxMCPServer")

# Create a state file for restoring state when MCP service restarts
CURRENT_DOC_FILE = os.path.join(tempfile.gettempdir(), "docx_mcp_current_doc.txt")

