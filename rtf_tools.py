from docx import Document
import fitz
documento = r"D:\28875-23-Anexo\Anexo\Eq01\Anexo_28875-23_Eq01.pdf"
    
# Create a Word document
def obs_interval_to_toc(pagina1=10167, x1=110, y1=200, pagina2=10175, x2=295, y2=390):
    doc = Document()
    table = doc.add_table(rows=2, cols=1, style="Table Grid")

    # Fill table cells with content
    for row in table.rows:
        for cell in row.cells:
            cell.text = "Cell Content"

    # Save the document
    doc.save("example.docx")
    
def extract_text_with_margins(pdf_path=documento, page_numbers=range(10167, 10176), top_margin=115, bottom_margin=815):
    """
    Extract text from specific pages of a PDF with specified top and bottom margins.
    
    :param pdf_path: Path to the PDF file
    :param page_numbers: List of page numbers to extract text from (0-based index)
    :param top_margin: Top margin in points (default is 50)
    :param bottom_margin: Bottom margin in points (default is 50)
    :return: Dictionary of extracted text by page number
    """
    # Open the PDF file
    doc = fitz.open(pdf_path)
    
    extracted_text = {}

    for page_num in page_numbers:
        # Get the page object
        page = doc.load_page(page_num)
        
        # Get the page dimensions
        page_width = page.rect.width
        page_height = page.rect.height
        
        # Define the clipping box (margins) in absolute coordinates
        top_clip = top_margin
        bottom_clip = bottom_margin
        
        # Extract text within the defined margins using a rectangular clip
        clip_rect = fitz.Rect(0, top_clip, page_width, bottom_clip)
        
        # Extract text from the clipped region
        text = page.get_text("text", clip=clip_rect)
        
        # Store the extracted text
        extracted_text[page_num] = text
    
    # Close the PDF document
    doc.close()
    
    return extracted_text
texto = extract_text_with_margins()
print(extract_text_with_margins())