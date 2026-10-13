"""
main.py

Phase 1 prototype: no Gmail yet. Drop PDFs into the "input/" folder using
the naming convention:

    <RegNo>_<Student Name>_<Department>.pdf

Run this script. For each PDF in input/:
    1. Parse metadata from the filename (filename_parser.py)
    2. Extract the text (pdf_extractor.py)
    3. Run it through Pangram's AI + plagiarism check (pangram_check.py)
    4. Append a row to output/master_log.xlsx (excel_logger.py)
    5. Move the PDF to input/processed/ (or input/needs_review/ on error)
       so it isn't re-processed next run.

Every PDF gets a log row, including ones that fail, so nothing silently
disappears. One bad PDF never stops the rest of the batch.

Usage:
    pip install -r requirements.txt
    cp .env.example .env        # then fill in your real Pangram API key
    python main.py
"""

from __future__ import annotations

import shutil
import sys
from datetime import datetime
from pathlib import Path

from excel_logger import LogWriteError, append_row, print_log_table
from filename_parser import FilenameParseError, StudentSubmission, parse_filename
from pdf_extractor import PDFExtractionError, extract_text
from pangram_check import CheckResult, check_paper

# Anchor all paths to this file's folder so the script works no matter
# which directory it is launched from.
BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "input"
PROCESSED_DIR = INPUT_DIR / "processed"
NEEDS_REVIEW_DIR = INPUT_DIR / "needs_review"
OUTPUT_LOG = BASE_DIR / "output" / "master_log.xlsx"


def _log(
    pdf_name: str,
    submission: StudentSubmission | None,
    status: str,
    result: CheckResult | None = None,
) -> None:
    """Write one log row. `submission`/`result` are None when unavailable."""
    append_row(
        OUTPUT_LOG,
        student_name=submission.student_name if submission else "",
        reg_no=submission.reg_no if submission else "",
        department=submission.department if submission else "",
        source_file=pdf_name,
        ai_percent=result.ai_percent if result else None,
        human_percent=result.human_percent if result else None,
        ai_tier=result.ai_tier if result else None,
        plagiarism_detected=result.plagiarism_detected if result else None,
        plagiarism_match_count=result.plagiarism_match_count if result else None,
        status=status,
    )


def process_one(pdf_path: Path) -> None:
    print(f"\n--- Processing {pdf_path.name} ---")

    # Step 1: parse filename -> student metadata
    try:
        submission = parse_filename(pdf_path)
    except FilenameParseError as e:
        print(f"  [FLAGGED] Could not parse filename: {e}")
        _log(pdf_path.name, None, "Needs Review (bad filename)")
        _move_to(pdf_path, NEEDS_REVIEW_DIR)
        return

    for warning in submission.warnings:
        print(f"  [WARNING] {warning}")

    # Step 2: extract PDF text
    try:
        text = extract_text(pdf_path)
    except PDFExtractionError as e:
        print(f"  [FLAGGED] Could not extract text: {e}")
        _log(pdf_path.name, submission, "Needs Review (unreadable PDF)")
        _move_to(pdf_path, NEEDS_REVIEW_DIR)
        return
    except Exception as e:  # noqa: BLE001 - corrupt/encrypted PDFs raise library errors
        print(f"  [FLAGGED] PDF could not be opened: {type(e).__name__}: {e}")
        _log(
            pdf_path.name,
            submission,
            f"Needs Review (PDF error: {type(e).__name__})",
        )
        _move_to(pdf_path, NEEDS_REVIEW_DIR)
        return

    # Step 3: run Pangram check
    try:
        result = check_paper(text)
    except Exception as e:  # noqa: BLE001 - surface any API/key error clearly
        print(f"  [ERROR] Pangram check failed: {e}")
        _log(pdf_path.name, submission, f"Error: {e}"[:250])
        _move_to(pdf_path, NEEDS_REVIEW_DIR)
        return

    # Step 4: log success. If the log can't be written, LogWriteError
    # propagates and stops the run; the PDF stays in input/ for next time.
    try:
        _log(pdf_path.name, submission, "Processed", result)
    except LogWriteError:
        print(
            "  [!] The Pangram check succeeded but the result could not be "
            f"logged. Result: AI={result.ai_percent} Tier={result.ai_tier} "
            f"Plagiarism={result.plagiarism_detected} "
            f"({result.plagiarism_match_count} matches)"
        )
        raise

    print(
        f"  [OK] {submission.student_name} ({submission.reg_no}): "
        f"AI={result.ai_percent} Plagiarism={result.plagiarism_detected}"
    )
    _move_to(pdf_path, PROCESSED_DIR)


def _unique_target(target_dir: Path, name: str) -> Path:
    """Return a path in target_dir that doesn't exist yet (never overwrite)."""
    target = target_dir / name
    if not target.exists():
        return target
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    stem, suffix = Path(name).stem, Path(name).suffix
    target = target_dir / f"{stem}__{stamp}{suffix}"
    counter = 2
    while target.exists():
        target = target_dir / f"{stem}__{stamp}-{counter}{suffix}"
        counter += 1
    return target


def _move_to(pdf_path: Path, target_dir: Path) -> None:
    """Move the PDF; on failure warn instead of crashing (already logged)."""
    try:
        target_dir.mkdir(parents=True, exist_ok=True)
        shutil.move(str(pdf_path), str(_unique_target(target_dir, pdf_path.name)))
    except OSError as e:
        print(
            f"  [WARNING] Logged, but could not move '{pdf_path.name}' to "
            f"'{target_dir}': {e}. It will be processed again next run."
        )


def main() -> None:
    INPUT_DIR.mkdir(exist_ok=True)
    OUTPUT_LOG.parent.mkdir(exist_ok=True)

    pdfs = sorted(
        p
        for p in INPUT_DIR.iterdir()
        if p.is_file() and p.suffix.lower() == ".pdf"
    )

    if not pdfs:
        print(
            f"No PDFs found in '{INPUT_DIR}'. Drop some in using the "
            f"naming convention '<RegNo>_<Name>_<Department>.pdf' and "
            f"run again."
        )
        sys.exit(0)

    print(f"Found {len(pdfs)} PDF(s) to process.")
    for pdf_path in pdfs:
        try:
            process_one(pdf_path)
        except LogWriteError as e:
            # The log is unavailable, so continuing would lose results.
            print(f"\n[STOPPED] {e}")
            sys.exit(1)
        except Exception as e:  # noqa: BLE001 - one bad file must not stop the batch
            print(f"  [ERROR] Unexpected failure on {pdf_path.name}: {type(e).__name__}: {e}")
            try:
                _log(
                    pdf_path.name,
                    None,
                    f"Error: unexpected {type(e).__name__}"[:250],
                )
            except LogWriteError as log_err:
                print(f"\n[STOPPED] {log_err}")
                sys.exit(1)
            _move_to(pdf_path, NEEDS_REVIEW_DIR)

    print(f"\nDone. Log: {OUTPUT_LOG}\n")
    print_log_table(OUTPUT_LOG)


if __name__ == "__main__":
    main()
