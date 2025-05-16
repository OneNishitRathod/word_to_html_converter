import mammoth
import html
import re

def convert_docx_to_html(docx_path, font_family=None, font_size=None):
    """
    Convert a Word document to HTML using mammoth with optional font styling.
    
    Args:
        docx_path (str): Path to the Word document file.
        font_family (str, optional): Font family to apply to the HTML content.
        font_size (str, optional): Font size to apply to the HTML content.
        
    Returns:
        tuple: (html_content, raw_html_content)
            - html_content: Raw HTML content for code display (not escaped)
            - raw_html_content: Raw HTML content for rendering
    """
    with open(docx_path, "rb") as docx_file:
        # Convert the document
        result = mammoth.convert_to_html(docx_file)
        html_content = result.value
        
        # Apply font styling if specified
        if font_family or font_size:
            # Create style attributes
            style_attrs = []
            if font_family:
                style_attrs.append(f"font-family: {font_family}")
            if font_size:
                style_attrs.append(f"font-size: {font_size}")
            
            style_str = "; ".join(style_attrs)
            
            # Wrap the content in a div with the specified style
            html_content = f'<div style="{style_str}">{html_content}</div>'
        
        # Get any messages from the conversion
        messages = result.messages
        
        # Return raw HTML for both code display and rendering
        return html_content, html_content

def get_plain_text_from_html(html_content):
    """
    Extract plain text from HTML content by removing all HTML tags.
    
    Args:
        html_content (str): HTML content.
        
    Returns:
        str: Plain text content.
    """
    # Remove all HTML tags
    text = re.sub(r'<[^>]+>', '', html_content)
    # Replace &nbsp; with spaces
    text = text.replace('&nbsp;', ' ')
    # Replace other common HTML entities
    text = html.unescape(text)
    # Normalize whitespace
    text = re.sub(r'\s+', ' ', text).strip()
    
    return text