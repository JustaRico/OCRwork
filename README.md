# Questionnaire Answer Extractor

A Python application that processes scanned or photographed filled-in Dutch questionnaires and extracts the answers per question, outputting them to an Excel file.

## Features

- **Multi-format input**: Accepts PDF files, images (PNG, JPG, JPEG, TIFF, BMP, GIF)
- **OCR Processing**: Uses Tesseract OCR to extract text from scans/photos
- **Dutch Language Support**: Optimized for Dutch language questionnaires
- **Excel Output**: Exports extracted answers to Excel (.xlsx) or CSV format
- **Batch Processing**: Can process multiple pages/images at once
- **Smart Question Matching**: Identifies questions and associates answers using pattern matching

## Requirements

- Python 3.8+
- Tesseract OCR (system package)
- Python packages: pymupdf, pillow, pytesseract, pandas, openpyxl, opencv-python-headless

## Installation

### 1. Install System Dependencies

**Ubuntu/Debian:**
```bash
sudo apt-get update
sudo apt-get install -y tesseract-ocr
```

**macOS:**
```bash
brew install tesseract
```

**Windows:**
Download and install from: https://github.com/UB-Mannheim/tesseract/wiki

### 2. Install Python Packages

```bash
pip install pymupdf pillow pytesseract pandas openpyxl opencv-python-headless
```

Or simply:
```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage

```bash
python questionnaire_extractor.py <input_file> [options]
```

### Examples

**Process a single image:**
```bash
python questionnaire_extractor.py scan_page1.png
```

**Process a PDF with multiple pages:**
```bash
python questionnaire_extractor.py questionnaire_scan.pdf -o answers.xlsx
```

**Process multiple images:**
```bash
python questionnaire_extractor.py image1.jpg image2.png image3.tiff --output results.xlsx
```

**Output both Excel and CSV:**
```bash
python questionnaire_extractor.py scan.pdf --csv
```

**Verbose mode:**
```bash
python questionnaire_extractor.py scan.pdf --verbose
```

### Command Line Options

```
positional arguments:
  input_files           Input image or PDF files (scans/photos of filled questionnaires)

optional arguments:
  -h, --help            show this help message and exit
  -o OUTPUT, --output OUTPUT
                        Output Excel file path (default: extracted_answers.xlsx)
  --csv                 Also save results as CSV
  --verbose             Show detailed processing information
```

## Output Format

The application generates an Excel file with two columns:

| Vraag (Question) | Antwoord (Answer) |
|------------------|-------------------|
| Respondentnummer | 12345 |
| Hoe oud ben je? | 25 |
| Waar woon je? | Amsterdam |

## How It Works

1. **Input Processing**: Converts PDF pages or loads images
2. **Image Preprocessing**: Enhances images for better OCR accuracy (grayscale, contrast enhancement)
3. **OCR Text Extraction**: Uses Tesseract to extract text from images
4. **Question-Answer Parsing**: Matches extracted text against known questions and identifies corresponding answers
5. **Output Generation**: Creates Excel/CSV file with question-answer pairs

## Tips for Best Results

1. **Image Quality**: Use high-resolution scans (300 DPI or higher)
2. **Lighting**: Ensure even lighting when photographing questionnaires
3. **Orientation**: Keep questionnaires straight and avoid rotation
4. **Handwriting**: Clear, legible handwriting produces better results
5. **Contrast**: Good contrast between writing and paper improves accuracy

## Questionnaire Structure

This tool is designed for the Dutch questionnaire in this repository which includes:

- Demographic questions (age, education, location)
- News consumption habits
- Content preferences
- Multiple choice and open-ended questions

## Limitations

- **Handwriting Recognition**: Accuracy depends on handwriting clarity
- **Complex Layouts**: May struggle with heavily formatted documents
- **Checkbox Detection**: Basic checkbox detection; manual verification recommended
- **Language**: Currently optimized for Dutch; other languages may require configuration changes

## Troubleshooting

### Poor OCR Results
- Increase image resolution
- Improve lighting conditions
- Ensure text is horizontal/straight
- Try preprocessing images manually

### Missing Answers
- Check if handwriting is too faint
- Verify question text matches template
- Review verbose output for debugging

### Tesseract Not Found
- Ensure Tesseract is installed and in system PATH
- On Windows, you may need to specify Tesseract path

## License

MIT License

## Contributing

Contributions welcome! Please feel free to submit issues and pull requests.
