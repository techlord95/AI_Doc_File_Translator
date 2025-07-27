from pdf2docx import Converter # Use pdf2docx
from io import FileIO
from os import path
import logging
import aspose.pdf as apdf

logger = logging.getLogger(__name__)

# Rename function to reflect the library used
def convert_pdf_to_docx_pdf2docx(pdf_path, word_path, start=None, end=None):
    """Converts a PDF file to DOCX using pdf2docx, optionally specifying pages.
    Args:
        pdf_path (str): Path to the input PDF file.
        word_path (str): Path to save the output DOCX file.
        start (int, optional): The first page to convert (0-indexed). Defaults to None (start from beginning).
        end (int, optional): The page number to stop converting before (0-indexed). 
                             Defaults to None (convert to the end). 
                             Use end=1 to convert only the first page (page 0).
    """
    logger.info(f"Attempting PDF to DOCX conversion using pdf2docx: {pdf_path} -> {word_path} (Requested Pages: start={start}, end={end})")
    try:
        cv = Converter(pdf_path)
        
        # Call cv.convert with specific arguments ONLY if they are not None
        if start is not None and end is not None:
            logger.debug(f"Calling cv.convert with start={start}, end={end}")
            cv.convert(word_path, start=start, end=end)
        elif start is not None:
            logger.debug(f"Calling cv.convert with start={start}")
            cv.convert(word_path, start=start)
        elif end is not None:
            # Workaround for potential pdf2docx issue: set start=0 if only end is provided
            logger.debug(f"Calling cv.convert with start=0, end={end} (start was None)")
            cv.convert(word_path, start=0, end=end) 
        else: # Both start and end are None
            logger.debug("Calling cv.convert with no page args (convert all)")
            cv.convert(word_path)
            
        cv.close()
        logger.info(f"pdf2docx conversion successful: {word_path}")
    except Exception as e:
        logger.error(f"pdf2docx conversion failed for {pdf_path}. Error: {e}")
        try:
            cv.close()
        except: pass
        raise

def convert_pdf_with_aspose(pdf_path, word_path, start=None, end=None):
    """Converts a PDF file to DOCX using Aspose.PDF.
    
    Args:
        pdf_path (str): Path to the input PDF file
        word_path (str): Path to save the output DOCX file
        start (int, optional): Not used in this implementation
        end (int, optional): Not used in this implementation
    """
    logger.info(f"Attempting PDF to DOCX conversion using Aspose: {pdf_path} -> {word_path}")
    
    try:
        # Load the document
        document = apdf.Document(pdf_path)
        
        # Configure save options
        save_options = apdf.DocSaveOptions()
        save_options.format = apdf.DocSaveOptions.DocFormat.DOC_X
        
        # Save the document
        document.save(word_path, save_options)
        logger.info(f"Aspose conversion successful: {word_path}")
        
    except Exception as e:
        logger.error(f"Aspose PDF conversion failed for {pdf_path}. Error: {e}")
        raise
