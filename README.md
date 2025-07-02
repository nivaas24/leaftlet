# PDF Reconstruction Project

A comprehensive Python toolkit for analyzing PDF structure, extracting content to editable formats, making modifications, and reconstructing PDFs with the same layout.

## Features

- **PDF Analysis**: Extract text, images, fonts, and layout information
- **Content Extraction**: Convert PDF content to Word documents for easy editing
- **Programmatic Editing**: Modify text, addresses, phone numbers, and other content
- **PDF Reconstruction**: Convert modified content back to PDF format
- **Multi-column Support**: Handle complex layouts with multiple columns
- **Image Handling**: Extract and preserve images during the conversion process

## Project Structure

```
pdf_reconstruction_project/
│
├── requirements.txt              # Python dependencies
├── main_workflow.py             # Complete workflow demonstration
├── README.md                    # This documentation
│
├── samples/                     # Sample PDF files
│   ├── sample.pdf              # Basic sample PDF
│   └── enhanced_sample.pdf     # Enhanced PDF with columns and images
│
├── src/                        # Source code modules
│   ├── pdf_analyzer.py         # PDF structure analysis
│   ├── pdf_to_word.py          # PDF to Word conversion
│   ├── word_to_pdf.py          # Word to PDF conversion
│   ├── edit_word_document.py   # Word document editing tools
│   ├── generate_sample_pdf.py  # Basic PDF generator
│   └── create_enhanced_pdf.py  # Enhanced PDF generator
│
└── output/                     # Generated files
    ├── analysis.json           # PDF structure analysis
    ├── images/                 # Extracted images
    ├── *.docx                  # Word documents
    └── *.pdf                   # Reconstructed PDFs
```

## Installation

1. **Clone or create the project directory**
2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

## Dependencies

- `PyPDF2` - PDF manipulation
- `pdfplumber` - Advanced PDF text extraction
- `reportlab` - PDF generation
- `Pillow` - Image processing
- `pymupdf` - PDF rendering and analysis
- `python-docx` - Word document handling
- `docx2pdf` - Word to PDF conversion
- `openpyxl` - Excel support (optional)

## Usage

### Quick Start - Complete Workflow

Run the complete demonstration:

```bash
python main_workflow.py
```

This will:
1. Generate sample PDFs (if they don't exist)
2. Analyze PDF structure
3. Extract content to Word documents
4. Show how to make modifications
5. Reconstruct the PDF

### Step-by-Step Usage

#### 1. Analyze PDF Structure

```python
from src.pdf_analyzer import PDFAnalyzer

with PDFAnalyzer("samples/enhanced_sample.pdf") as analyzer:
    # Save detailed analysis
    analyzer.save_analysis("output/analysis.json")
    
    # Extract images
    analyzer.extract_images_to_folder("output/images")
```

#### 2. Convert PDF to Word

```python
from src.pdf_to_word import PDFToWordConverter

converter = PDFToWordConverter("samples/enhanced_sample.pdf", "output/extracted.docx")
converter.convert_to_word()
```

#### 3. Edit Word Document

**Manual editing**: Open the generated `.docx` file in Microsoft Word and make changes.

**Programmatic editing**:
```python
from src.edit_word_document import WordDocumentEditor

editor = WordDocumentEditor("output/extracted.docx")

# Replace text
editor.replace_text("Lorem ipsum", "Updated content")

# Replace contact information
editor.replace_address("456 New Street, Updated City, NY 10001")
editor.replace_phone("(555) 987-6543")
editor.replace_email("new@email.com")

# Add new content
editor.add_header("New Section", level=2)
editor.add_paragraph("This is additional content.")

# Save changes
editor.save("output/modified.docx")
```

#### 4. Convert Back to PDF

```python
from src.word_to_pdf import WordToPDFConverter

converter = WordToPDFConverter("output/modified.docx", "output/final.pdf")
converter.convert_to_pdf()
```

### Demo Scripts

#### Enhanced PDF Demo

Create and process a complex PDF with multiple columns and images:

```bash
python src/edit_word_document.py
```

#### Basic Workflow Demo

Process a simple PDF:

```bash
python src/pdf_analyzer.py
```

## Advanced Features

### PDF Analysis Output

The `analysis.json` file contains:

```json
{
  "metadata": {
    "title": "Document Title",
    "author": "Author Name",
    "creation_date": "...",
    ...
  },
  "pages": [
    {
      "page_number": 0,
      "width": 612.0,
      "height": 792.0,
      "text_elements": [
        {
          "text": "Sample text",
          "x0": 100.0,
          "y0": 50.0,
          "x1": 200.0,
          "y1": 70.0,
          "fontname": "Helvetica",
          "fontsize": 12.0,
          "page_number": 0
        }
      ],
      "images": [...],
      "columns": [...]
    }
  ],
  "fonts": [...],
  "total_pages": 1
}
```

### Supported Modifications

- **Text replacement**: Any text content
- **Address updates**: Using regex pattern matching
- **Phone number changes**: Format: (XXX) XXX-XXXX
- **Email updates**: Standard email format
- **Content addition**: New paragraphs and headers
- **Image preservation**: Images are maintained during conversion

### Custom Text Patterns

You can extend the `WordDocumentEditor` class to handle custom patterns:

```python
def replace_custom_pattern(self, pattern: str, replacement: str):
    """Replace text using regex patterns"""
    for paragraph in self.doc.paragraphs:
        if re.search(pattern, paragraph.text):
            for run in paragraph.runs:
                run.text = re.sub(pattern, replacement, run.text)
```

## Limitations

1. **Complex Layouts**: Very complex layouts may not convert perfectly
2. **Font Preservation**: Some fonts may change during conversion
3. **Image Quality**: Image quality might be affected during extraction/insertion
4. **Microsoft Word Dependency**: `docx2pdf` requires Microsoft Word to be installed
5. **Table Formatting**: Complex table structures may need manual adjustment

## Troubleshooting

### Common Issues

1. **docx2pdf Error**: Ensure Microsoft Word is installed on Windows
2. **Image Extraction Failed**: Some PDF images may be embedded in complex ways
3. **Font Issues**: Default fonts will be used if original fonts are not available
4. **Layout Changes**: Simple layouts work best; complex designs may require manual adjustment

### Alternative Approaches

For Linux/Mac systems without Microsoft Word:
```bash
# Use LibreOffice for conversion
pip install python-docx2pdf
# Or use online conversion services
```

## Examples

### Example 1: Update Company Information

```python
# Load document
editor = WordDocumentEditor("output/extracted.docx")

# Update all company information
editor.replace_text("Old Company Name", "New Company Name")
editor.replace_address("123 New Business Blvd, Suite 100, City, ST 12345")
editor.replace_phone("(555) 123-NEWCO")
editor.replace_email("info@newcompany.com")

# Save and convert
editor.save("output/updated.docx")

# Convert back to PDF
converter = WordToPDFConverter("output/updated.docx", "output/updated.pdf")
converter.convert_to_pdf()
```

### Example 2: Batch Processing Multiple PDFs

```python
import os
from pathlib import Path

pdf_folder = "input_pdfs/"
output_folder = "output_pdfs/"

for pdf_file in Path(pdf_folder).glob("*.pdf"):
    # Extract to Word
    word_file = f"temp_{pdf_file.stem}.docx"
    converter = PDFToWordConverter(str(pdf_file), word_file)
    converter.convert_to_word()
    
    # Make modifications
    editor = WordDocumentEditor(word_file)
    editor.replace_text("OLD_TEXT", "NEW_TEXT")
    editor.save()
    
    # Convert back to PDF
    output_pdf = f"{output_folder}/modified_{pdf_file.name}"
    pdf_converter = WordToPDFConverter(word_file, output_pdf)
    pdf_converter.convert_to_pdf()
    
    # Clean up temporary file
    os.remove(word_file)
```

## Contributing

1. Fork the repository
2. Create a feature branch
3. Add your improvements
4. Test with various PDF types
5. Submit a pull request

## License

This project is open-source. Feel free to modify and distribute according to your needs.

## Support

For issues or questions:
1. Check the troubleshooting section
2. Review the example code
3. Test with the provided sample PDFs
4. Create an issue with sample files for debugging
