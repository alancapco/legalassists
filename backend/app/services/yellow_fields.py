"""Extract and fill yellow-highlighted fields in corporate Word templates.

Corporate authors mark placeholders in Word with yellow highlight. This module:
1. Walks paragraphs and table cells
2. Collects contiguous yellow runs as a single field occurrence
3. Derives a stable field key from the placeholder text
4. Fills values and clears the highlight on export
"""

from __future__ import annotations

import hashlib
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable

from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph


YELLOW = WD_COLOR_INDEX.YELLOW

# Explicit key forms: {{party_name}}, [[PARTY_NAME]], [PARTY NAME]
KEY_PATTERNS = (
    re.compile(r"^\{\{\s*([a-zA-Z0-9_.-]+)\s*\}\}$"),
    re.compile(r"^\[\[\s*([a-zA-Z0-9_.-]+)\s*\]\]$"),
    re.compile(r"^\[\s*([A-Za-z0-9_ .-]+)\s*\]$"),
)


def slugify_key(text: str) -> str:
    cleaned = text.strip().lower()
    cleaned = re.sub(r"[\{\}\[\]<>]", "", cleaned)
    cleaned = re.sub(r"[^a-z0-9]+", "_", cleaned)
    cleaned = cleaned.strip("_")
    return cleaned or "field"


def derive_field_key(placeholder: str) -> tuple[str, str]:
    """Return (key, human_label) from yellow placeholder text."""
    raw = " ".join(placeholder.split()).strip()
    for pattern in KEY_PATTERNS:
        match = pattern.match(raw)
        if match:
            key = slugify_key(match.group(1))
            label = match.group(1).replace("_", " ").strip().title()
            return key, label
    key = slugify_key(raw)
    label = raw if raw else key.replace("_", " ").title()
    return key, label


def _is_yellow_run(run) -> bool:
    try:
        if run.font.highlight_color == YELLOW:
            return True
    except Exception:
        pass
    # Fallback: inspect raw XML highlight value
    rpr = run._element.find(qn("w:rPr"))
    if rpr is None:
        return False
    highlight = rpr.find(qn("w:highlight"))
    if highlight is None:
        return False
    return (highlight.get(qn("w:val")) or "").lower() == "yellow"


def _iter_paragraphs(document: Document) -> Iterable[Paragraph]:
    """Yield paragraphs in true top-to-bottom XML document order (body + tables interleaved)."""

    def _walk_container(parent_elm, parent_obj) -> Iterable[Paragraph]:
        for child in parent_elm.iterchildren():
            if child.tag == qn("w:p"):
                yield Paragraph(child, parent_obj)
            elif child.tag == qn("w:tbl"):
                table = Table(child, parent_obj)
                yield from _iter_table_paragraphs(table)

    yield from _walk_container(document.element.body, document)


def _iter_table_paragraphs(table: Table) -> Iterable[Paragraph]:
    for row in table.rows:
        seen_tcs: set[int] = set()
        for cell in row.cells:
            tc_id = id(cell._tc)
            if tc_id in seen_tcs:
                continue
            seen_tcs.add(tc_id)
            for child in cell._tc.iterchildren():
                if child.tag == qn("w:p"):
                    yield Paragraph(child, cell)
                elif child.tag == qn("w:tbl"):
                    yield from _iter_table_paragraphs(Table(child, cell))


@dataclass
class FieldOccurrence:
    key: str
    label: str
    placeholder: str
    paragraph_index: int
    run_start: int
    run_end: int  # exclusive
    context: str


@dataclass
class ScannedTemplate:
    fields: dict[str, dict] = field(default_factory=dict)
    occurrences: list[FieldOccurrence] = field(default_factory=list)
    preview_text: str = ""
    checksum: str = ""


def file_checksum(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(65536), b""):
            digest.update(chunk)
    return digest.hexdigest()


PLACEHOLDER_FIND = re.compile(
    r"\{\{[^{}]+\}\}|\[\[[^\[\]]+\]\]|\[[^\[\]]{1,80}\]"
)

# Inline choice options / optional clause prefixes that use [brackets] in legal templates
# but are NOT fillable field placeholders.
_OPTION_BRACKET_WORDS = frozenset(
    {
        "advance",
        "arrears",
        "do",
        "do not",
        "binding",
        "non-binding",
        "week",
        "month",
        "year",
        "an individual residing at",
        "an individual with a mailing address at",
        "as the milestone fees in section 3",
        "x",
    }
)


def is_valid_placeholder_token(token: str) -> bool:
    """Return True if a matched [bracket] / {{mustache}} token is a genuine fillable placeholder."""
    raw = (token or "").strip()
    if not raw:
        return False
    if raw.startswith("{{") or raw.startswith("[["):
        inner = raw[2:-2].strip()
        return bool(inner) and not re.fullmatch(r"[_\s$]+", inner)
    if raw.startswith("[") and raw.endswith("]"):
        inner = raw[1:-1].strip()
        if not inner:
            return False
        # Skip checkboxes [  ] or underscore-only tokens [__], [$___]
        if re.fullmatch(r"[_\s$]+", inner):
            return False
        # Skip optional clause brackets that contain underscore blanks, choice slashes, or full sentences
        if "___" in inner or " / " in inner or ". " in inner:
            return False
        # Skip optional clause brackets starting with punctuation like [, except that...]
        if inner[0] in ",;:.":
            return False
        lowered = " ".join(inner.lower().split())
        if lowered in _OPTION_BRACKET_WORDS:
            return False
        if lowered.startswith(
            (
                "or until ",
                "as follows:",
                "approved in advance",
                "total fees shall ",
                "not to exceed ",
                "the employer ",
            )
        ):
            return False
        return True
    return False


def ensure_placeholder_highlights(path: Path) -> int:
    """Yellow-highlight [bracket] / {{mustache}} placeholders so they become fillable fields.

    Used when a user uploads a plain .docx as a template. Returns how many placeholders
    were highlighted (0 if the document already used yellow fields only).
    """
    document = Document(str(path))
    highlighted = 0

    for paragraph in _iter_paragraphs(document):
        text = paragraph.text or ""
        matches = [
            m
            for m in PLACEHOLDER_FIND.finditer(text)
            if is_valid_placeholder_token(m.group(0))
        ]
        if not matches:
            continue
        # Skip rebuild when every match is already covered by yellow runs
        already = scan_paragraph_yellow_placeholders(paragraph)
        if already and len(already) >= len(matches):
            continue
        # Preserve nothing fancy — rebuild runs with yellow on placeholders
        p_el = paragraph._p
        for child in list(p_el):
            if child.tag == qn("w:r"):
                p_el.remove(child)
        cursor = 0
        for match in matches:
            if match.start() > cursor:
                paragraph.add_run(text[cursor : match.start()])
            run = paragraph.add_run(match.group(0))
            run.font.highlight_color = YELLOW
            highlighted += 1
            cursor = match.end()
        if cursor < len(text):
            paragraph.add_run(text[cursor:])

    if highlighted:
        document.save(str(path))
    return highlighted


def scan_paragraph_yellow_placeholders(paragraph: Paragraph) -> list[str]:
    found: list[str] = []
    runs = paragraph.runs
    i = 0
    while i < len(runs):
        if not _is_yellow_run(runs[i]):
            i += 1
            continue
        chunks: list[str] = []
        while i < len(runs) and _is_yellow_run(runs[i]):
            chunks.append(runs[i].text or "")
            i += 1
        placeholder = "".join(chunks).strip()
        if placeholder:
            found.append(placeholder)
    return found


def scan_docx(path: Path) -> ScannedTemplate:
    document = Document(str(path))
    result = ScannedTemplate(checksum=file_checksum(path))
    preview_parts: list[str] = []
    prev_non_empty: str = ""

    for p_index, paragraph in enumerate(_iter_paragraphs(document)):
        text = paragraph.text.strip()
        if text:
            preview_parts.append(text)

        runs = paragraph.runs
        i = 0
        while i < len(runs):
            if not _is_yellow_run(runs[i]):
                i += 1
                continue
            start = i
            chunks: list[str] = []
            while i < len(runs) and _is_yellow_run(runs[i]):
                chunks.append(runs[i].text or "")
                i += 1
            placeholder = "".join(chunks)
            if not placeholder.strip():
                continue
            key, label = derive_field_key(placeholder)
            # If the paragraph only contains the placeholder itself (e.g. a table value cell
            # or a standalone blank line), include the preceding label/prompt paragraph for context.
            if prev_non_empty and (text == placeholder.strip() or len(text) <= len(placeholder.strip()) + 12):
                ctx_str = f"{prev_non_empty} → {text}"[:240]
            else:
                ctx_str = text[:240]
            occurrence = FieldOccurrence(
                key=key,
                label=label,
                placeholder=placeholder.strip(),
                paragraph_index=p_index,
                run_start=start,
                run_end=i,
                context=ctx_str,
            )
            result.occurrences.append(occurrence)
            bucket = result.fields.setdefault(
                key,
                {
                    "key": key,
                    "label": label,
                    "placeholder": placeholder.strip(),
                    "occurrence_count": 0,
                    "required": True,
                    "field_type": "short_text",
                    "sample_context": ctx_str or None,
                },
            )
            bucket["occurrence_count"] += 1

        if text:
            prev_non_empty = text

    result.preview_text = "\n\n".join(preview_parts[:60])
    return result


def fill_docx(
    source: Path,
    destination: Path,
    values: dict[str, str],
    *,
    clear_highlight: bool = True,
) -> list[str]:
    """Fill yellow fields by key. Returns list of still-missing required keys."""
    document = Document(str(source))
    filled_keys: set[str] = set()
    seen_keys: set[str] = set()

    for paragraph in _iter_paragraphs(document):
        runs = paragraph.runs
        i = 0
        while i < len(runs):
            if not _is_yellow_run(runs[i]):
                i += 1
                continue
            start = i
            chunks: list[str] = []
            while i < len(runs) and _is_yellow_run(runs[i]):
                chunks.append(runs[i].text or "")
                i += 1
            placeholder = "".join(chunks)
            if not placeholder.strip():
                continue
            key, _label = derive_field_key(placeholder)
            seen_keys.add(key)
            value = values.get(key)
            if value is None or value == "":
                continue
            # Put full replacement text on first run; clear the rest
            runs[start].text = value
            for run in runs[start + 1 : i]:
                run.text = ""
            if clear_highlight:
                for run in runs[start:i]:
                    run.font.highlight_color = None
            filled_keys.add(key)

    destination.parent.mkdir(parents=True, exist_ok=True)
    document.save(str(destination))
    # python-docx strips theme/settings — restore a full package for EigenPal
    from app.services.docx_package import enrich_docx_for_editor

    enrich_docx_for_editor(destination)
    return sorted(seen_keys - filled_keys)
