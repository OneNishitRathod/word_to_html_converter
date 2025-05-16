import mammoth
import html

def convert_docx_to_html(docx_path):
    """
    Convert a Word document to HTML using mammoth.
    
    Args:
        docx_path (str): Path to the Word document file.
        
    Returns:
        str: The HTML content and raw HTML code.
    """
    with open(docx_path, "rb") as docx_file:
        # Convert the document
        result = mammoth.convert_to_html(docx_file)
        html_content = result.value
        
        # Get any messages from the conversion
        messages = result.messages
        
        # Escape the HTML for displaying as code
        html_content_escaped = html.escape(html_content)
        
        return html_content_escaped
