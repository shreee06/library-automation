"""
email_parser.py

Best-effort extraction of student details from an email body. The library's
final email format isn't settled, so this never guesses: anything it can't
find is returned blank and the row is flagged "Needs Review".

Looks for:
    SRN     -- strict PES format only, e.g. "pes1pg23ca133" / "PES2UG21CS045"
    Name    -- only when labeled, e.g. "Name: Priya Sharma" / "my name is ..."
    Mentor  -- when labeled, e.g. "Mentor: Prof X" / "mentor prof X"
    (Name/mentor are capped at 4 words and cut at words like "and"/"your".)

Like filename_parser.py, this is the only file that needs to change once
the real email format is decided.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

# Strict PES SRN: PES + digit + UG/PG + 2-digit year + dept code + number,
# e.g. PES1PG23CA133, PES2UG21CS045. Anything else -> "Needs Review".
_SRN_RE = re.compile(r"\bPES[0-9](?:UG|PG)\d{2}[A-Z]{2,3}\d{3,4}\b", re.IGNORECASE)

_NAME_RE = re.compile(
    r"\b(?:student\s+)?name\b[ \t]*(?:[:\-–=]|\bis\b)[ \t]*([^\r\n]+)",
    re.IGNORECASE,
)
_MENTOR_RE = re.compile(
    r"\bmentor\b[ \t]*(?:[:\-–=]|\bis\b)?[ \t]*([^\r\n]+)",
    re.IGNORECASE,
)

MAX_NAME_WORDS = 4
# A captured name/mentor is cut off at the first of these words.
_STOP_WORDS = {
    "and", "your", "my", "our", "with", "from", "for", "to", "of", "is",
    "at", "in", "on", "the", "this", "that", "which", "who", "email", "srn",
}


@dataclass
class EmailStudentInfo:
    srn: str = ""
    student_name: str = ""
    mentor: str = ""
    needs_review: bool = True
    review_reason: str = ""


def _clean(value: str) -> str:
    """Trim a captured name/mentor: cut at a stop word, cap the word count."""
    words: list[str] = []
    for word in value.split():
        if word.lower().strip(".,;:") in _STOP_WORDS:
            break
        words.append(word)
        if len(words) == MAX_NAME_WORDS:
            break
    return " ".join(words).strip(".,;:-– ").strip()


def parse_email_body(body: str) -> EmailStudentInfo:
    srns: list[str] = []
    for tok in _SRN_RE.findall(body or ""):
        if tok.lower() not in (s.lower() for s in srns):
            srns.append(tok)

    name_match = _NAME_RE.search(body or "")
    mentor_match = _MENTOR_RE.search(body or "")

    info = EmailStudentInfo(
        srn=srns[0] if len(srns) == 1 else "",
        student_name=_clean(name_match.group(1)) if name_match else "",
        mentor=_clean(mentor_match.group(1)) if mentor_match else "",
    )

    reasons: list[str] = []
    if not srns:
        reasons.append("No PES-format SRN found in body")
    elif len(srns) > 1:
        reasons.append("Multiple SRNs found: " + ", ".join(srns))
    if not info.student_name:
        reasons.append("No student name found in body")

    info.review_reason = "; ".join(reasons)
    info.needs_review = bool(reasons)
    return info


if __name__ == "__main__":
    for sample in [
        "pes1pg23ca133 my srn and mentor prof shishir N",
        "Name: Priya Sharma\nSRN: PES2UG21CS045\nMentor: Dr. Rao",
        "Hello, please find my report attached.",
    ]:
        print(parse_email_body(sample))
