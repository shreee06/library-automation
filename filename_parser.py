"""
filename_parser.py

Since we don't have library email access yet, student metadata (name,
registration number, department) comes from the PDF's FILENAME instead of
an email body. This enforces a simple, fixed naming convention so parsing
is reliable from day one:

    <RegNo>_<Student Name>_<Department>.pdf

Examples:
    21CS045_John_Doe_CSE.pdf
    21ME102_Priya_Sharma_MECH.pdf

If your library later standardizes on a web form or a fixed email subject
line instead, only this file needs to change -- everything downstream
(pangram_check.py, excel_logger.py) just consumes the dict this returns.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class StudentSubmission:
    reg_no: str
    student_name: str
    department: str
    source_filename: str
    warnings: list[str] = field(default_factory=list)


class FilenameParseError(ValueError):
    """Raised when a PDF's filename doesn't match the expected convention."""


def parse_filename(pdf_path: str | Path) -> StudentSubmission:
    """
    Parse "<RegNo>_<Name parts...>_<Department>.pdf" into a StudentSubmission.

    The name may contain underscores itself (e.g. multiple words), so we
    take the FIRST underscore-separated token as the reg no, the LAST as
    the department, and everything in between (joined with spaces) as the
    student's name.
    """
    path = Path(pdf_path)
    stem = path.stem  # filename without ".pdf"
    parts = stem.split("_")

    if len(parts) < 3:
        raise FilenameParseError(
            f"Filename '{path.name}' does not match "
            f"'<RegNo>_<Name>_<Department>.pdf' (need at least 3 "
            f"underscore-separated parts, got {len(parts)})."
        )

    reg_no = parts[0].strip()
    department = parts[-1].strip()
    student_name = " ".join(parts[1:-1]).strip()

    if not reg_no or not student_name or not department:
        raise FilenameParseError(
            f"Filename '{path.name}' parsed to an empty field "
            f"(reg_no={reg_no!r}, name={student_name!r}, dept={department!r})."
        )

    # No strict format yet: only warn (never fail) on unusual reg numbers.
    warnings: list[str] = []
    if not reg_no.isalnum():
        warnings.append(
            f"Registration number {reg_no!r} is not purely alphanumeric; "
            f"check the filename '{path.name}'."
        )

    return StudentSubmission(
        reg_no=reg_no,
        student_name=student_name,
        department=department,
        source_filename=path.name,
        warnings=warnings,
    )


if __name__ == "__main__":
    # Quick manual test: python filename_parser.py
    samples = ["21CS045_John_Doe_CSE.pdf", "21ME102_Priya_Sharma_MECH.pdf"]
    for s in samples:
        print(parse_filename(s))
