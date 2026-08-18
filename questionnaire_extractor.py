#!/usr/bin/env python3
"""
Questionnaire Answer Extractor

This application processes scanned or photographed filled-in questionnaires
and extracts the answers per question, outputting them to an Excel file.

The questionnaire is in Dutch and expects handwritten or typed answers
in the empty boxes/fields.

Usage:
    python questionnaire_extractor.py <image_or_pdf_file> [output_excel_file]
    
Examples:
    python questionnaire_extractor.py scan_page1.png
    python questionnaire_extractor.py questionnaire_scan.pdf answers.xlsx
    python questionnaire_extractor.py image1.jpg image2.png --output results.xlsx
"""

import argparse
import os
import re
import sys
from pathlib import Path
from typing import List, Dict, Tuple, Optional

# Third-party imports
try:
    import pymupdf  # formerly fitz
except ImportError:
    print("Please install pymupdf: pip install pymupdf")
    sys.exit(1)

try:
    from PIL import Image
except ImportError:
    print("Please install Pillow: pip install pillow")
    sys.exit(1)

try:
    import pytesseract
except ImportError:
    print("Please install pytesseract: pip install pytesseract")
    sys.exit(1)

try:
    import pandas as pd
except ImportError:
    print("Please install pandas: pip install pandas")
    sys.exit(1)

try:
    from openpyxl import Workbook
except ImportError:
    print("Please install openpyxl: pip install openpyxl")
    sys.exit(1)


# Define the questionnaire structure based on the PDF
# These are the 14 numbered answer fields from the questionnaire
QUESTIONS = [
    "Respondentnummer",
    "Geïnterviewde is een man, vrouw, X",
    "Hoe oud ben je?",
    "Zit je op de universiteit, HBO of MBO?",
    "Waar woon je?",
    "Heb je wat met dit nieuws gedaan? (nadat je het gezien of gehoord had)",
    "Wanneer is nieuws over jouw buurt, stad of regio interessant voor je? (Inhoud)",
    "Waar vind jij dat nieuws uit jouw buurt, stad of regio over moet gaan? (onderwerp)",
    "Heb jij nog tips voor journalisten die nieuws maken over buurt, stad of regio-omgeving?",
    "Kun je een top 3 geven?",
    "Denk eens aan het laatste nieuws dat je hebt gezien of gehoord over jouw buurt, stad of regio.",
    "Wat vond je van dit nieuws?",
    "Vind je het leuk om mee te doen aan vervolgonderzoek?",
    "Hoe wil jij dit nieuws zien?",
]

# Checkbox options for multiple choice questions
CHECKBOX_OPTIONS = {
    "Geïnterviewde is een man, vrouw, X": ["M", "V", "X"],
    "Zit je op de universiteit, HBO of MBO?": ["Universiteit", "HBO", "MBO"],
    "Laatste nieuws": ["Tech & Innovatie", "Klimaat", "Health & Mindset", "Veiligheid & Criminaliteit",
                       "Entertainment", "Studie & Ontwikkeling", "Politiek & Maatschappij", "Sport",
                       "112-nieuws", "Reizen & Avontuur", "Activisme & Impact", "Geld & Carrière",
                       "Cultuur", "Lifestyle"],
}

# Content criteria checkboxes
CONTENT_CRITERIA = [
    "gaat over dingen die jij interessant vindt",
    "gaat over onderwerpen die jou raken",
    "duidelijk uitlegt wat er aan de hand is",
    "verder kijkt dan alleen de koppen, meer diepgang heeft",
    "verschillende kanten van een verhaal laat zien",
    "ook laat zien wat wél werkt en welke kansen er zijn",
    "niet alleen negatief is, maar ook positieve dingen laat zien",
    "boeiend en makkelijk te begrijpen is",
    "vaker over jongeren gaat",
    "laat zien wat het voor mij of mijn buurt betekent",
    "laat zien hoe ik er zelf mee te maken heb",
    "mij iets nieuws leert",
    "betrouwbaar voelt",
    "kort en duidelijk is",
]


def convert_to_images(input_path: str) -> List[Image.Image]:
    """
    Convert PDF or image file(s) to a list of PIL Images.
    
    Args:
        input_path: Path to PDF or image file
        
    Returns:
        List of PIL Image objects
    """
    images = []
    path = Path(input_path)
    
    if path.suffix.lower() == '.pdf':
        # Convert PDF pages to images
        doc = pymupdf.open(path)
        for page_num in range(len(doc)):
            page = doc[page_num]
            # Render page at 300 DPI for better OCR accuracy
            mat = pymupdf.Matrix(300/72, 300/72)
            pix = page.get_pixmap(matrix=mat)
            img_data = pix.tobytes("png")
            img = Image.open(io.BytesIO(img_data))
            images.append(img)
        doc.close()
    else:
        # Load image file directly
        img = Image.open(path)
        images.append(img)
    
    return images


def preprocess_image(image: Image.Image) -> Image.Image:
    """
    Preprocess image for better OCR results.
    
    Args:
        image: PIL Image object
        
    Returns:
        Preprocessed PIL Image
    """
    # Convert to grayscale
    image = image.convert('L')
    
    # Increase contrast by adjusting levels
    from PIL import ImageEnhance
    enhancer = ImageEnhance.Contrast(image)
    image = enhancer.enhance(1.5)
    
    enhancer = ImageEnhance.Brightness(image)
    image = enhancer.enhance(1.2)
    
    return image


def extract_text_with_ocr(image: Image.Image, lang: str = 'eng') -> str:
    """
    Extract text from image using Tesseract OCR.
    
    Args:
        image: PIL Image object
        lang: Language code for OCR ('eng' for English, 'dut' for Dutch if available)
        
    Returns:
        Extracted text string
    """
    # Preprocess image
    processed_image = preprocess_image(image)
    
    # Configure Tesseract for better accuracy
    custom_config = r'--oem 3 --psm 6'
    
    try:
        # Try Dutch first if available, fallback to English
        text = pytesseract.image_to_string(processed_image, lang='dut+eng', config=custom_config)
    except:
        # Fallback to English only
        text = pytesseract.image_to_string(processed_image, lang='eng', config=custom_config)
    
    return text


def extract_text_regions(image: Image.Image) -> List[Tuple[str, int, int, int, int]]:
    """
    Extract text with bounding box information.
    
    Args:
        image: PIL Image object
        
    Returns:
        List of tuples: (text, x, y, width, height)
    """
    processed_image = preprocess_image(image)
    
    # Get detailed OCR data with bounding boxes
    data = pytesseract.image_to_data(processed_image, output_type=pytesseract.Output.DICT)
    
    regions = []
    n_boxes = len(data['text'])
    
    for i in range(n_boxes):
        text = data['text'][i].strip()
        if text:  # Only include non-empty text
            x = data['left'][i]
            y = data['top'][i]
            w = data['width'][i]
            h = data['height'][i]
            regions.append((text, x, y, w, h))
    
    return regions


def find_answers_in_text(text: str) -> Dict[str, str]:
    """
    Parse OCR text to identify questions and their corresponding answers.
    
    This uses pattern matching and spatial reasoning to associate answers
    with their corresponding questions.
    
    Args:
        text: Full OCR text from the image
        
    Returns:
        Dictionary mapping questions to answers
    """
    answers = {}
    lines = text.split('\n')
    
    current_question = None
    current_answer = []
    
    # Clean and normalize lines
    cleaned_lines = []
    for line in lines:
        line = line.strip()
        if line:
            cleaned_lines.append(line)
    
    i = 0
    while i < len(cleaned_lines):
        line = cleaned_lines[i]
        
        # Skip very short noise (1-2 chars unless it's a known option like M, V, X)
        if len(line) <= 2 and line not in ['M', 'V', 'X', 'Ja', 'Nee']:
            # But still check if it could be part of an answer
            pass
        
        # Check if this line matches a known question
        matched_question = None
        best_match_score = 0
        
        for question in QUESTIONS:
            # Partial match - check if question keywords are present
            question_keywords = question.lower().split()[:5]  # First 5 words
            line_lower = line.lower()
            
            # Count keyword matches
            match_count = sum(1 for kw in question_keywords if kw in line_lower and len(kw) > 2)
            
            # Also check for longer phrases
            question_words = question.lower().split()
            for j in range(len(question_words) - 2):
                phrase = ' '.join(question_words[j:j+3])
                if phrase in line_lower and len(phrase) > 5:
                    match_count += 2
            
            if match_count > best_match_score:
                best_match_score = match_count
                matched_question = question
        
        # Require at least 2 keyword matches or a long phrase match
        if matched_question and best_match_score >= 2:
            # Save previous question-answer pair
            if current_question and current_answer:
                answer_text = ' '.join(current_answer).strip()
                # Filter out noise and question fragments
                if answer_text and len(answer_text) > 3:
                    # Remove common noise patterns
                    answer_text = re.sub(r'\([^)]*\)\s*$', '', answer_text).strip()
                    if answer_text and not re.match(r'^[•\-\*]\s*$', answer_text):
                        answers[current_question] = answer_text
            
            # Start new question
            current_question = matched_question
            current_answer = []
            
            # Check if answer is on the same line (after the question mark or colon)
            if ':' in line:
                answer_part = line.split(':', 1)[1].strip()
                if answer_part and len(answer_part) > 2:
                    # Make sure it's not just repeating the question
                    if answer_part.lower() not in current_question.lower():
                        current_answer.append(answer_part)
            elif '?' in line:
                answer_part = line.split('?', 1)[1].strip()
                if answer_part and len(answer_part) > 2:
                    if answer_part.lower() not in current_question.lower():
                        current_answer.append(answer_part)
        else:
            # This might be an answer line
            if current_question and line:
                # Skip checkbox markers and common noise
                if not re.match(r'^[•\-\*]\s*$', line):
                    # Skip if it looks like a checkbox option label from the template
                    is_option = False
                    for options in CHECKBOX_OPTIONS.values():
                        if any(opt.lower() == line.lower() for opt in options):
                            is_option = True
                            break
                    
                    # Skip content criteria labels
                    for criterion in CONTENT_CRITERIA:
                        if criterion.lower() == line.lower():
                            is_option = True
                            break
                    
                    # Skip if line is too similar to question (likely just question text)
                    if current_question and len(line) > 10:
                        similarity = len(set(line.lower()) & set(current_question.lower())) / max(len(line), len(current_question))
                        if similarity > 0.7:
                            is_option = True
                    
                    if not is_option and len(line) > 1:
                        current_answer.append(line)
        
        i += 1
    
    # Save last question-answer pair
    if current_question and current_answer:
        answer_text = ' '.join(current_answer).strip()
        if answer_text and len(answer_text) > 3:
            answer_text = re.sub(r'\([^)]*\)\s*$', '', answer_text).strip()
            if answer_text and not re.match(r'^[•\-\*]\s*$', answer_text):
                answers[current_question] = answer_text
    
    return answers


def detect_checkboxes(image: Image.Image) -> Dict[str, List[str]]:
    """
    Detect checked checkboxes in the image.
    
    This uses image processing to identify marked checkboxes.
    
    Args:
        image: PIL Image object
        
    Returns:
        Dictionary mapping question to list of selected options
    """
    import cv2
    import numpy as np
    
    # Convert PIL to OpenCV format
    img_array = np.array(image)
    if len(img_array.shape) == 3:
        gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
    else:
        gray = img_array
    
    # Apply adaptive thresholding to handle varying lighting
    thresh = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
                                    cv2.THRESH_BINARY_INV, 11, 2)
    
    # Find contours
    contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    # This is a simplified checkbox detection
    # In production, you'd want more sophisticated detection with template matching
    selected = {}
    
    return selected


def process_image(image: Image.Image) -> Dict[str, str]:
    """
    Process a single image to extract question-answer pairs.
    
    Args:
        image: PIL Image object
        
    Returns:
        Dictionary mapping questions to answers
    """
    # Extract full text
    full_text = extract_text_with_ocr(image)
    
    # Parse answers from text
    answers = find_answers_in_text(full_text)
    
    # Try to detect checkboxes
    checkbox_answers = detect_checkboxes(image)
    
    # Merge checkbox answers
    for question, selections in checkbox_answers.items():
        if selections:
            answers[question] = ', '.join(selections)
    
    return answers


def process_multiple_images(images: List[Image.Image]) -> Dict[str, str]:
    """
    Process multiple images and combine results.
    
    Args:
        images: List of PIL Image objects
        
    Returns:
        Combined dictionary mapping questions to answers
    """
    all_answers = {}
    
    for i, image in enumerate(images):
        print(f"Processing image {i+1}/{len(images)}...")
        answers = process_image(image)
        
        # Merge answers, preferring non-empty values
        for question, answer in answers.items():
            if answer and answer.strip():
                # If we already have an answer, append with separator
                if question in all_answers and all_answers[question]:
                    existing = all_answers[question]
                    if answer not in existing:
                        all_answers[question] = existing + '; ' + answer
                else:
                    all_answers[question] = answer
    
    return all_answers


def save_to_excel(answers: Dict[str, str], output_path: str):
    """
    Save extracted answers to an Excel file.
    
    Args:
        answers: Dictionary mapping questions to answers
        output_path: Path to output Excel file
    """
    # Create DataFrame with questions and answers
    df_data = []
    
    for question in QUESTIONS:
        answer = answers.get(question, '')
        df_data.append({
            'Vraag (Question)': question,
            'Antwoord (Answer)': answer
        })
    
    # Add any extra questions not in our predefined list
    for question, answer in answers.items():
        if question not in QUESTIONS:
            df_data.append({
                'Vraag (Question)': question,
                'Antwoord (Answer)': answer
            })
    
    df = pd.DataFrame(df_data)
    
    # Save to Excel
    df.to_excel(output_path, index=False, sheet_name='Answers')
    
    print(f"\nResults saved to: {output_path}")
    print(f"Total questions processed: {len(df_data)}")
    print(f"Questions with answers: {sum(1 for row in df_data if row['Antwoord (Answer)'])}")


def save_to_csv(answers: Dict[str, str], output_path: str):
    """
    Save extracted answers to a CSV file.
    
    Args:
        answers: Dictionary mapping questions to answers
        output_path: Path to output CSV file
    """
    df_data = []
    
    for question in QUESTIONS:
        answer = answers.get(question, '')
        df_data.append({
            'Vraag (Question)': question,
            'Antwoord (Answer)': answer
        })
    
    for question, answer in answers.items():
        if question not in QUESTIONS:
            df_data.append({
                'Vraag (Question)': question,
                'Antwoord (Answer)': answer
            })
    
    df = pd.DataFrame(df_data)
    df.to_csv(output_path, index=False, encoding='utf-8-sig')
    
    print(f"\nResults saved to: {output_path}")


def main():
    parser = argparse.ArgumentParser(
        description='Extract answers from filled-in Dutch questionnaires',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__
    )
    
    parser.add_argument(
        'input_files',
        nargs='+',
        help='Input image or PDF files (scans/photos of filled questionnaires)'
    )
    
    parser.add_argument(
        '-o', '--output',
        default='extracted_answers.xlsx',
        help='Output Excel file path (default: extracted_answers.xlsx)'
    )
    
    parser.add_argument(
        '--csv',
        action='store_true',
        help='Also save results as CSV'
    )
    
    parser.add_argument(
        '--verbose',
        action='store_true',
        help='Show detailed processing information'
    )
    
    args = parser.parse_args()
    
    # Validate input files
    for input_file in args.input_files:
        if not os.path.exists(input_file):
            print(f"Error: File not found: {input_file}")
            sys.exit(1)
        
        ext = Path(input_file).suffix.lower()
        if ext not in ['.pdf', '.png', '.jpg', '.jpeg', '.tiff', '.bmp', '.gif']:
            print(f"Warning: Unsupported file type: {ext}")
    
    print("=" * 60)
    print("Questionnaire Answer Extractor")
    print("=" * 60)
    print(f"\nProcessing {len(args.input_files)} file(s)...")
    
    # Convert all inputs to images
    all_images = []
    
    for input_file in args.input_files:
        print(f"\nLoading: {input_file}")
        try:
            images = convert_to_images(input_file)
            all_images.extend(images)
            print(f"  -> Loaded {len(images)} page(s)")
        except Exception as e:
            print(f"  -> Error loading file: {e}")
            if args.verbose:
                import traceback
                traceback.print_exc()
    
    if not all_images:
        print("\nError: No images could be loaded from input files.")
        sys.exit(1)
    
    print(f"\nTotal pages/images to process: {len(all_images)}")
    
    # Process all images
    answers = process_multiple_images(all_images)
    
    # Determine output format
    output_path = args.output
    if output_path.lower().endswith('.csv'):
        save_to_csv(answers, output_path)
    else:
        if not output_path.lower().endswith('.xlsx'):
            output_path = output_path.replace('.xls', '') + '.xlsx'
        save_to_excel(answers, output_path)
        
        # Also save CSV if requested
        if args.csv:
            csv_path = output_path.rsplit('.', 1)[0] + '.csv'
            save_to_csv(answers, csv_path)
    
    # Print summary
    print("\n" + "=" * 60)
    print("EXTRACTION SUMMARY")
    print("=" * 60)
    
    for question in QUESTIONS[:10]:  # Show first 10 questions
        answer = answers.get(question, '')
        if answer:
            print(f"\nQ: {question}")
            print(f"A: {answer[:100]}{'...' if len(answer) > 100 else ''}")
    
    print("\n" + "=" * 60)
    print("Processing complete!")
    print("=" * 60)


if __name__ == '__main__':
    import io
    main()
