"""
pdf_extractor.py

Pulls plain text out of a PDF so it can be sent to Pangram for checking.
"""

from __future__ import annotations

from pathlib import Path

import pdfplumber


class PDFExtractionError(ValueError):
    """Raised when a PDF has no extractable text (e.g. it's a scan/image)."""


def extract_text(pdf_path: str | Path) -> str:
    """
    Extract all text from a PDF, page by page, joined with blank lines.

    Raises PDFExtractionError if the PDF yields no usable text at all
    (e.g. it's a scanned image with no OCR layer) -- that's the signal
    main.py uses to flag a submission for manual review instead of
    silently sending empty text to Pangram.
    """
    path = Path(pdf_path)
    pages_text: list[str] = []

    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if text.strip():
                pages_text.append(text)

    full_text = "\n\n".join(pages_text).strip()

    if not full_text:
        raise PDFExtractionError(
            f"No extractable text found in '{path.name}'. "
            f"It may be a scanned/image-only PDF that needs OCR first."
        )

    return full_text


if __name__ == "__main__":
    # Quick manual test: python pdf_extractor.py path/to/file.pdf
    import sys

    if len(sys.argv) != 2:
        print("Usage: python pdf_extractor.py <path-to-pdf>")
        sys.exit(1)

    text = extract_text(sys.argv[1])
    print(f"Extracted {len(text)} characters.")
    print(text[:500] + ("..." if len(text) > 500 else ""))
