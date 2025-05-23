from flask import Flask, render_template, request, flash, redirect, url_for, jsonify
import os
from werkzeug.utils import secure_filename
from utils.converter import convert_docx_to_html, analyze_document_structure

# Flask application setup
app = Flask(__name__)
app.config['SECRET_KEY'] = 'word2html-converter-app'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024  # 32MB max file size
app.config['ALLOWED_EXTENSIONS'] = {'docx', 'doc'}

# Available font options
FONT_FAMILIES = [
    {"value": "", "name": "Default"},
    {"value": "Arial, sans-serif", "name": "Arial"},
    {"value": "Times New Roman, serif", "name": "Times New Roman"},
    {"value": "Calibri, sans-serif", "name": "Calibri"},
    {"value": "Verdana, sans-serif", "name": "Verdana"},
    {"value": "Georgia, serif", "name": "Georgia"},
    {"value": "Courier New, monospace", "name": "Courier New"},
    {"value": "Helvetica, sans-serif", "name": "Helvetica"},
    {"value": "Tahoma, sans-serif", "name": "Tahoma"}
]

FONT_SIZES = [
    {"value": "", "name": "Default"},
    {"value": "10px", "name": "10px"},
    {"value": "12px", "name": "12px"},
    {"value": "14px", "name": "14px"},
    {"value": "16px", "name": "16px"},
    {"value": "18px", "name": "18px"},
    {"value": "20px", "name": "20px"},
    {"value": "24px", "name": "24px"},
    {"value": "28px", "name": "28px"}
]

# Image resize options
IMAGE_SIZES = [
    {"value": "600", "name": "Small (600px)"},
    {"value": "800", "name": "Medium (800px)"},
    {"value": "1200", "name": "Large (1200px)"},
    {"value": "1600", "name": "Extra Large (1600px)"},
    {"value": "2000", "name": "Original Size (2000px)"}
]

# Create upload folder if it doesn't exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Check if the post request has the file part
        if 'document' not in request.files:
            flash('No file part')
            return redirect(request.url)
        
        file = request.files['document']
        
        # If user does not select file, browser also
        # submit an empty part without filename
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        
        # Get formatting options from form
        font_family = request.form.get('font_family', '')
        font_size = request.form.get('font_size', '')
        max_image_width = int(request.form.get('image_size', '800'))
        
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            
            try:
                # Analyze document structure first
                doc_analysis = analyze_document_structure(filepath)
                
                # Convert document with enhanced options
                html_content_escaped, raw_html_content = convert_docx_to_html(
                    filepath,
                    font_family=font_family,
                    font_size=font_size,
                    max_image_width=max_image_width
                )
                
                # Remove uploaded file after conversion
                os.remove(filepath)
                
                return render_template(
                    'result.html',
                    html_content=html_content_escaped,
                    raw_html_content=raw_html_content,
                    doc_analysis=doc_analysis,
                    original_filename=file.filename
                )
            except Exception as e:
                flash(f'Error converting document: {str(e)}')
                # Clean up file on error
                if os.path.exists(filepath):
                    os.remove(filepath)
                return redirect(request.url)
        else:
            flash('File type not allowed. Please upload a .docx or .doc file.')
            return redirect(request.url)
    
    return render_template(
        'index.html',
        font_families=FONT_FAMILIES,
        font_sizes=FONT_SIZES,
        image_sizes=IMAGE_SIZES
    )

@app.route('/api/analyze', methods=['POST'])
def analyze_document():
    """API endpoint to analyze document structure without full conversion."""
    if 'document' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    
    file = request.files['document']
    
    if file.filename == '' or not allowed_file(file.filename):
        return jsonify({'error': 'Invalid file'}), 400
    
    filename = secure_filename(file.filename)
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(filepath)
    
    try:
        analysis = analyze_document_structure(filepath)
        os.remove(filepath)  # Clean up
        return jsonify(analysis)
    except Exception as e:
        if os.path.exists(filepath):
            os.remove(filepath)
        return jsonify({'error': str(e)}), 500

@app.errorhandler(413)
def too_large(e):
    flash('File too large. Please upload a file smaller than 32MB.')
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)