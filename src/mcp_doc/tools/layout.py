"""Layout and formatting tools"""

from typing import Optional
from docx.shared import Cm

from ..config import logger
from ..processor import processor


def add_page_break() -> str:
    """
    Add page break
    """
    try:
        if not processor.current_document:
            return "No document is open"
        
        processor.current_document.add_page_break()
        
        return "Page break added"
    except Exception as e:
        error_msg = f"Failed to add page break: {str(e)}"
        logger.error(error_msg)
        return error_msg


def set_page_margins(
    top: Optional[float] = None,
    bottom: Optional[float] = None,
    left: Optional[float] = None,
    right: Optional[float] = None
) -> str:
    """
    Set page margins
    
    Parameters:
    - top: Top margin (cm)
    - bottom: Bottom margin (cm)
    - left: Left margin (cm)
    - right: Right margin (cm)
    """
    try:
        if not processor.current_document:
            return "No document is open"
        
        doc = processor.current_document
        
        # Get current section (default to use first section)
        section = doc.sections[0]
        
        # Set page margins
        if top is not None:
            section.top_margin = Cm(top)
        if bottom is not None:
            section.bottom_margin = Cm(bottom)
        if left is not None:
            section.left_margin = Cm(left)
        if right is not None:
            section.right_margin = Cm(right)
        
        return "Page margins set"
    except Exception as e:
        error_msg = f"Failed to set page margins: {str(e)}"
        logger.error(error_msg)
        return error_msg

