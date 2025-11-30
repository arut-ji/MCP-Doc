"""DocxProcessor class for processing Word documents"""

from typing import Dict, Optional, Any
from docx import Document

from .config import CURRENT_DOC_FILE, logger
import os

# Type aliases for better type checking
# Document from docx has incomplete type stubs
DocumentType = Any  # docx.Document has incomplete type stubs


class DocxProcessor:
    """Class for processing Docx documents, implementing various document operations"""
    
    def __init__(self) -> None:
        self.documents: Dict[str, DocumentType] = {}  # Store opened documents
        self.current_document: Optional[DocumentType] = None
        self.current_file_path: Optional[str] = None
        
        # Try to load current document from state file
        self._load_current_document()
    
    def _load_current_document(self) -> bool:
        """Load current document from state file"""
        if not os.path.exists(CURRENT_DOC_FILE):
            return False
        
        try:
            with open(CURRENT_DOC_FILE, 'r', encoding='utf-8') as f:
                file_path = f.read().strip()
            
            if file_path and os.path.exists(file_path):
                try:
                    self.current_file_path = file_path
                    self.current_document = Document(file_path)
                    self.documents[file_path] = self.current_document
                    return True
                except Exception as e:
                    logger.error(f"Failed to load document at {file_path}: {e}")
                    # Delete invalid state file to prevent future loading attempts
                    try:
                        os.remove(CURRENT_DOC_FILE)
                        logger.info(f"Removed invalid state file pointing to {file_path}")
                    except Exception as e_remove:
                        logger.error(f"Failed to remove state file: {e_remove}")
            else:
                # Delete invalid state file if path is empty or file doesn't exist
                try:
                    os.remove(CURRENT_DOC_FILE)
                    logger.info("Removed invalid state file with non-existent document path")
                except Exception as e_remove:
                    logger.error(f"Failed to remove state file: {e_remove}")
        except Exception as e:
            logger.error(f"Failed to load current document: {e}")
            # Delete corrupted state file
            try:
                os.remove(CURRENT_DOC_FILE)
                logger.info("Removed corrupted state file")
            except Exception as e_remove:
                logger.error(f"Failed to remove state file: {e_remove}")
        
        return False
    
    def _save_current_document(self) -> bool:
        """Save current document path to state file"""
        if not self.current_file_path:
            return False
        
        try:
            with open(CURRENT_DOC_FILE, 'w', encoding='utf-8') as f:
                f.write(self.current_file_path)
            return True
        except Exception as e:
            logger.error(f"Failed to save current document path: {e}")
        
        return False
    
    def save_state(self) -> None:
        """Save processor state"""
        # Save current document
        if self.current_document and self.current_file_path:
            try:
                self.current_document.save(self.current_file_path)
                self._save_current_document()
            except Exception as e:
                logger.error(f"Failed to save current document: {e}")
    
    def load_state(self) -> None:
        """Load processor state"""
        self._load_current_document()


# Create global processor instance
processor = DocxProcessor()

