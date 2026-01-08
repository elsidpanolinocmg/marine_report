# ocr_utils.py
import re
import pytesseract
from PIL import Image

# configure tesseract path here if needed (app can override too)
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

def clean_ocr_text(text: str) -> str:
    """Fix broken OCR text: hyphens, newlines, spacing"""
    if not text:
        return ""
    # join hyphen-split words across lines
    text = re.sub(r"(\w)-\n(\w)", r"\1\2", text)
    # collapse many newlines into single
    text = re.sub(r"\n+", "\n", text)
    # single newlines (line wraps) -> space; preserve paragraph breaks (double newlines)
    text = re.sub(r"(?<!\n)\n(?!\n)", " ", text)
    return text.strip()

def ocr_image_crop(pil_image: Image.Image, bbox: tuple, psm: int = 6) -> str:
    """
    Run tesseract OCR on a cropped region (bbox in image coordinates).
    bbox: (x1, y1, x2, y2)
    """
    crop = pil_image.crop(bbox)
    try:
        raw = pytesseract.image_to_string(crop, config=f"--psm {psm}")
    except Exception as e:
        raw = f"[ocr error: {e}]"
    return clean_ocr_text(raw)


def clean_paragraphs(lines):
    """
    Merge OCR lines into paragraphs using:
    - indentation (leading spaces)
    - punctuation at end of line
    - short line heuristic
    """
    paragraphs = []
    current = []
    lengths = [len(l.strip()) for l in lines if l.strip()]
    avg_len = sum(lengths) / len(lengths) if lengths else 0

    def flush():
        nonlocal current
        if current:
            paragraphs.append(" ".join(current).strip())
            current = []

    for line in lines:
        l = line.rstrip()
        if not l.strip():
            flush()
            continue

        is_indented = l.startswith("  ")  # 2+ spaces
        if current:
            if is_indented or current[-1].endswith((".", "!", "?", ":")) or len(l.strip()) < avg_len * 0.4:
                flush()
        current.append(l.strip())

    flush()
    return paragraphs
