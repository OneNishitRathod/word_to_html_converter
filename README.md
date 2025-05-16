# Word to HTML Converter

A simple web application that converts Microsoft Word documents (.docx, .doc) to HTML code.

## Features

- Upload Word documents and convert them to HTML
- View the raw HTML code
- Preview the HTML rendering
- Copy HTML code to clipboard with a single click
- Clean, responsive user interface

## Project Structure

```
word_to_html_converter/
├── app.py              # Main Flask application entry point
├── templates/          # HTML templates
│   ├── index.html      # Main page template
│   └── result.html     # Result page template
├── static/             # Static files (CSS, JS)
│   ├── css/
│   │   └── style.css   # Custom styling
│   └── js/
│       └── script.js   # Frontend scripting
├── utils/              # Utility functions
│   └── converter.py    # Word to HTML conversion logic
├── requirements.txt    # Project dependencies
└── README.md           # Project documentation
```

## Requirements

- Python 3.7+
- Flask
- Mammoth
- Werkzeug

## Installation

1. Clone the repository:
   ```
   git clone https://github.com/yourusername/word-to-html-converter.git
   cd word-to-html-converter
   ```

2. Create a virtual environment:
   ```
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. Install dependencies:
   ```
   pip install -r requirements.txt
   ```

## Usage

1. Run the application:
   ```
   python app.py
   ```

2. Open your browser and navigate to `http://127.0.0.1:5000`

3. Upload a Word document and click "Convert to HTML"

4. View the HTML code and preview, copy the code as needed

## How It Works

The application uses the Mammoth library to convert Word documents to HTML. When a document is uploaded, the Flask application processes it through Mammoth and displays both the raw HTML code and a preview of how it would render in a browser.

## License

This project is licensed under the MIT License - see the LICENSE file for details.