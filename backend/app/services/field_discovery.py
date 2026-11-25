"""Discover fillable blanks in uploaded Word templates and yellow-highlight them.

Handles:
- Existing [brackets] / {{mustache}} / [[keys]]
- 2-column key-value form tables & multi-column deliverables/milestone/rate/invention tables
- Underscore blanks (____, $____, ____%) using defined-term lookahead, inline label lookbehind,
  local clause rules, and preceding paragraph context
- Underlined whitespace / tab form blanks
- Party-scoped signature block fields (never turning Signature:/By: lines into text fields)
- AI Semantic Refinement (renaming any remaining generic [Additional Term] placeholders using LLM clause understanding)
- AI Precedent Variable Discovery (turning hardcoded dates, party names, fees, and governing law in precedent contracts into yellow [Fields])
"""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from loguru import logger

from app.services.ai import copilot
from app.services.yellow_fields import (
    PLACEHOLDER_FIND,
    YELLOW,
    _is_yellow_run,
    _iter_paragraphs,
    ensure_placeholder_highlights,
    is_valid_placeholder_token,
    scan_docx,
    slugify_key,
)

# Underscore / currency / date blanks common in contract templates
BLANK_FIND = re.compile(
    r"\$_{2,}|_{3,}|\b20xx\b|\bXXXX\b|\bxx/xx/xxxx\b|\bMM/DD/YYYY\b",
    re.IGNORECASE,
)

# Map 2-column form table prompts (left cell) to clean semantic field labels (right cell)
TABLE_PROMPT_LABELS: list[tuple[re.Pattern[str], str]] = [
    (re.compile(r"master agreement", re.I), "Master Agreement Details"),
    (re.compile(r"provider\s*\(full legal name\)", re.I), "Provider Name"),
    (re.compile(r"client\s*\(full legal name\)", re.I), "Client Name"),
    (re.compile(r"provider and client\s*\(full legal names\)", re.I), "Provider and Client Names"),
    (re.compile(r"project name", re.I), "Project Name"),
    (re.compile(r"provider[’']?s project manager", re.I), "Provider Project Manager"),
    (re.compile(r"client[’']?s project manager", re.I), "Client Project Manager"),
    (re.compile(r"^in scope\b", re.I), "In Scope Services"),
    (re.compile(r"^out of scope\b", re.I), "Out of Scope"),
    (re.compile(r"^key dependencies", re.I), "Key Dependencies"),
    (re.compile(r"services:\s*areas on which the consultant", re.I), "Consulting Services Scope"),
    (re.compile(r"services\s*\(itemized description\)", re.I), "Services Description"),
    (re.compile(r"description of the services", re.I), "Services Description"),
    (re.compile(r"company[’']?s key contact", re.I), "Company Key Contact"),
    (re.compile(r"individual who will perform the services", re.I), "Consultant Individual"),
    (re.compile(r"^deliverables\b", re.I), "Deliverables"),
    (re.compile(r"^location of the services", re.I), "Services Location"),
    (re.compile(r"^start date,?\s*timeline and end date", re.I), "Project Timeline and Dates"),
    (re.compile(r"^total value of the services", re.I), "Total Value of Services"),
    (re.compile(r"^deposit or advance payment", re.I), "Deposit Amount"),
    (re.compile(r"^deposit payable on signing", re.I), "Deposit Amount"),
    (re.compile(r"^invoicing\b", re.I), "Invoicing Schedule"),
    (re.compile(r"^insurance the contractor must maintain", re.I), "Required Insurance Coverage"),
    (re.compile(r"^third-party materials and contractor materials", re.I), "Pre-Existing and Third-Party Materials"),
    (re.compile(r"^expenses and materials the client pays for", re.I), "Reimbursable Expenses and Materials"),
    (re.compile(r"^payment method and details", re.I), "Payment Method Details"),
    (re.compile(r"^end date of the term", re.I), "Term End Date"),
    (re.compile(r"^exclusivity during the term", re.I), "Exclusivity Terms"),
]

# Party role headers in signature blocks
PARTY_HEADER_RE = re.compile(
    r"^(?:(PROVIDER|CLIENT|COMPANY|CONSULTANT|CONTRACTOR|CUSTOMER|VENDOR|LICENSOR|LICENSEE|DISCLOSER|RECIPIENT|EMPLOYER|EMPLOYEE|CANDIDATE|ASSIGNOR|ASSIGNEE|PARTY\s*A|PARTY\s*B)\b"
    r"|(Provider|Client|Company|Consultant|Contractor|Customer|Vendor|Licensor|Licensee|Discloser|Recipient|Employer|Employee|Candidate|Assignor|Assignee|Party\s*A|Party\s*B)\s*:)",
    re.IGNORECASE,
)


def _is_signature_line(text: str) -> bool:
    """Return True if a line is a wet-ink / electronic signature line (Signature: ___ or By: ___)."""
    stripped = (text or "").strip()
    return bool(re.match(r"^(?:Signature|By)\s*:\s*[_\s]*$", stripped, re.IGNORECASE))


def _detect_party_header(text: str) -> str | None:
    stripped = (text or "").strip()
    match = PARTY_HEADER_RE.match(stripped)
    if not match:
        return None
    role = (match.group(1) or match.group(2) or "").strip()
    return role.title() if role else None


def _signature_field_label(line_text: str, party_role: str | None) -> str | None:
    """If line_text is a signature metadata line (Name:, Title:, Date signed:, etc.), return a party-scoped label."""
    stripped = (line_text or "").strip()
    prefix = f"{party_role} " if party_role else ""
    if re.match(r"^Name(?:\s+in\s+full)?\s*:", stripped, re.I):
        return f"{prefix}Signatory Name".strip()
    if re.match(r"^Title(?:\s*\([^)]*\))?\s*:", stripped, re.I):
        return f"{prefix}Signatory Title".strip()
    if re.match(r"^Date(?:\s+signed)?\s*:", stripped, re.I):
        return f"{prefix}Date Signed".strip()
    if re.match(r"^Entity\s+name(?:\s*\([^)]*\))?\s*:", stripped, re.I):
        return f"{prefix}Entity Name".strip()
    if re.match(r"^Personal\s+email(?:\s+for\s+notices)?\s*:", stripped, re.I):
        return f"{prefix}Personal Email".strip()
    if re.match(r"^Email(?:\s+for\s+notices)?\s*:", stripped, re.I):
        return f"{prefix}Notice Email".strip()
    return None


def _strip_template_instructions(document: Document) -> None:
    """Remove trailing 'Notes for completing this template' instruction pages."""
    body = document.element.body
    removing_notes = False
    for child in list(body.iterchildren()):
        if child.tag == qn("w:p"):
            text = "".join(child.itertext()).strip()
            if re.match(r"^Notes for completing this template\b", text, re.I):
                removing_notes = True
            if removing_notes:
                body.remove(child)


def _is_change_order_exhibit(text: str) -> bool:
    """Return True only for a 'Form of Change Order' exhibit (not substantive Exhibit A schedules)."""
    stripped = (text or "").strip().lower()
    return "form of change order" in stripped


def _set_cell_yellow_placeholder(cell, placeholder: str, *, prefix: str = "", suffix: str = "") -> None:
    """Write a yellow-highlighted placeholder into a table cell's first paragraph."""
    paragraphs = cell.paragraphs
    p = paragraphs[0] if paragraphs else cell.add_paragraph()
    p_el = p._p
    for child in list(p_el):
        if child.tag == qn("w:r"):
            p_el.remove(child)
    if prefix:
        p.add_run(prefix)
    run = p.add_run(placeholder)
    run.font.highlight_color = YELLOW
    if suffix:
        p.add_run(suffix)


def _set_cell_multi_placeholders(cell, parts: list[tuple[str, bool]]) -> None:
    """Write mixed plain/yellow runs into a table cell's first paragraph."""
    paragraphs = cell.paragraphs
    p = paragraphs[0] if paragraphs else cell.add_paragraph()
    p_el = p._p
    for child in list(p_el):
        if child.tag == qn("w:r"):
            p_el.remove(child)
    for text, is_yellow in parts:
        if not text:
            continue
        run = p.add_run(text)
        if is_yellow:
            run.font.highlight_color = YELLOW


def highlight_table_blanks(path: Path) -> int:
    """Populate 2-column form tables and multi-column schedule tables with semantic yellow fields."""
    document = Document(str(path))
    _strip_template_instructions(document)
    highlighted = 0
    in_change_order_exhibit = False

    for child in document.element.body.iterchildren():
        if child.tag == qn("w:p"):
            text = "".join(child.itertext()).strip()
            if _is_change_order_exhibit(text):
                in_change_order_exhibit = True
            continue
        if child.tag != qn("w:tbl") or in_change_order_exhibit:
            continue

        table = Table(child, document)
        if not table.rows:
            continue

        num_cols = len(table.columns)
        header_texts = [" ".join(c.text.split()).strip().lower() for c in table.rows[0].cells]

        # Case 1: 5-column Deliverables & Milestones schedule table
        if (
            num_cols == 5
            and any("deliverable" in h for h in header_texts)
            and any("acceptance" in h or "description" in h for h in header_texts)
        ):
            for r_idx, row in enumerate(table.rows[1:], start=1):
                cells = row.cells
                if all(not c.text.strip() for c in cells):
                    _set_cell_multi_placeholders(cells[0], [(str(r_idx), False)])
                    _set_cell_yellow_placeholder(cells[1], f"[Deliverable {r_idx} Name]")
                    _set_cell_yellow_placeholder(cells[2], f"[Deliverable {r_idx} Acceptance Criteria]")
                    _set_cell_yellow_placeholder(cells[3], f"[Deliverable {r_idx} Due Date]")
                    _set_cell_yellow_placeholder(cells[4], f"[Deliverable {r_idx} Fee]")
                    highlighted += 4
            continue

        # Case 2: 3-column Milestone / Deliverable / Due Date / Fee table
        if (
            num_cols == 3
            and any("milestone" in h or "deliverable" in h for h in header_texts)
            and any("due" in h or "date" in h for h in header_texts)
        ):
            for r_idx, row in enumerate(table.rows[1:], start=1):
                cells = row.cells
                if all(not c.text.strip() for c in cells):
                    _set_cell_yellow_placeholder(cells[0], f"[Milestone {r_idx} Deliverable]")
                    _set_cell_yellow_placeholder(cells[1], f"[Milestone {r_idx} Due Date]")
                    _set_cell_yellow_placeholder(cells[2], f"[Milestone {r_idx} Fee]")
                    highlighted += 3
            continue

        # Case 3: 3-column Prior Inventions table (Title, Date made, Patent/Application No.)
        if (
            num_cols == 3
            and any("title" in h or "description" in h for h in header_texts)
            and any("patent" in h or "application" in h for h in header_texts)
        ):
            # Populate Row 1 as an optional Prior Invention entry (instead of 4 redundant rows)
            if len(table.rows) > 1 and all(not c.text.strip() for c in table.rows[1].cells):
                cells = table.rows[1].cells
                _set_cell_yellow_placeholder(cells[0], "[Prior Invention Title]")
                _set_cell_yellow_placeholder(cells[1], "[Prior Invention Date]")
                _set_cell_yellow_placeholder(cells[2], "[Prior Invention Application Number]")
                highlighted += 3
            continue

        # Case 4: 2-column Key-Value form tables (Label in Col 0, Value in Col 1)
        if num_cols == 2:
            first_cell_upper = table.rows[0].cells[0].text.strip().upper()
            if first_cell_upper.startswith(
                ("PROVIDER", "CLIENT", "COMPANY", "CONSULTANT", "CONTRACTOR", "EMPLOYER", "EMPLOYEE")
            ):
                continue

            for row in table.rows:
                cells = row.cells
                if id(cells[0]._tc) == id(cells[1]._tc):
                    continue
                left_text = " ".join(cells[0].text.split()).strip()
                right_text = " ".join(cells[1].text.split()).strip()
                if not left_text:
                    continue
                if any(_is_yellow_run(r) for p in cells[1].paragraphs for r in p.runs):
                    continue

                if re.match(r"^No\.\s+Date\s*:\s*$", right_text, re.I):
                    _set_cell_multi_placeholders(
                        cells[1],
                        [
                            ("No. ", False),
                            ("[SOW Number]", True),
                            ("    Date: ", False),
                            ("[SOW Effective Date]", True),
                        ],
                    )
                    highlighted += 2
                    continue

                if re.match(r"^Start\s*:\s+End\s*:\s*$", right_text, re.I):
                    _set_cell_multi_placeholders(
                        cells[1],
                        [
                            ("Start: ", False),
                            ("[Start Date]", True),
                            ("    End: ", False),
                            ("[End Date]", True),
                        ],
                    )
                    highlighted += 2
                    continue

                # Empty right cell or "$" currency prompt in right cell
                if not right_text or right_text == "$":
                    if re.search(r"check\s+one", left_text, re.I):
                        continue
                    label = None
                    for pattern, mapped_label in TABLE_PROMPT_LABELS:
                        if pattern.search(left_text):
                            label = mapped_label
                            break
                    if not label:
                        base = re.split(r"[\(:]", left_text)[0].strip()
                        if base and len(base) <= 45:
                            label = base.title()
                    if label:
                        prefix = "$" if right_text == "$" else ""
                        _set_cell_yellow_placeholder(cells[1], f"[{label}]", prefix=prefix)
                        highlighted += 1

    if highlighted:
        document.save(str(path))
        logger.info("highlighted {} table form fields in {}", highlighted, path.name)
    return highlighted


def _clean_defined_term(term: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9 &/-]+", " ", term).strip()
    if cleaned.lower() in {"employer", "employee", "company", "consultant", "contractor", "provider", "client", "customer", "vendor"}:
        return f"{cleaned.title()} Name"
    if cleaned.lower() == "proposed collaboration":
        return "Proposed Collaboration Scope"
    return cleaned.title() if cleaned else "Field"


def _label_from_local_context(
    paragraph_text: str,
    blank: str,
    match_start: int,
    match_end: int,
    prev_paragraph_text: str = "",
) -> str:
    """Derive a semantic label for a blank using defined terms, inline labels, local window, and preceding paragraph."""
    if re.match(r"(?i)^(20xx|xxxx|xx/xx/xxxx|mm/dd/yyyy)$", blank):
        return "Effective Date"

    full = paragraph_text.strip()
    before = paragraph_text[max(0, match_start - 100) : match_start]
    after = paragraph_text[match_end : min(len(paragraph_text), match_end + 100)]

    # 1. If the entire paragraph is just underscores, use the preceding paragraph's prompt
    if re.fullmatch(r"[_\s$]+", full):
        prev = prev_paragraph_text.strip()
        if re.search(r"client shall provide|responsibilit", prev, re.I):
            return "Client Responsibilities"
        if re.search(r"based on the following assumptions|assumptions", prev, re.I):
            return "Project Assumptions"
        if re.search(r"special terms|changes the master agreement", prev, re.I):
            return "Special Terms"
        if re.search(r"scope|services", prev, re.I):
            return "Scope of Services"
        return "Additional Details"

    # 2. Immediate defined-term lookahead right after the blank:
    # e.g. `____ (the "Effective Date")`, `____ (the "Position")`, `____ (hereinafter referred to as "Employer")`
    def_after = re.match(
        r'^\s*\((?:the\s+|hereinafter\s+referred\s+to\s+as\s+)?[“"]([^”"]{2,40})[”"]\)',
        after,
        re.I,
    )
    if def_after:
        return _clean_defined_term(def_after.group(1))

    # 3. Preamble / Party definition lines: (1) ____ ... (the "Company" / "Employer" / "Client" / "Provider")
    party_role_match = re.search(r'\(the\s+[“"]([^”"]+)[”"]\)\s*[.;,]?\s*(?:and)?\s*$', full)
    if party_role_match:
        role = party_role_match.group(1).strip()
        if match_start < 15:
            return f"{role} Name"
        if re.search(r"\[?a\s+$", before, re.I):
            return f"{role} Entity Type"
        if re.search(r"laws of\s+$", before, re.I):
            return f"{role} Jurisdiction"
        if re.search(r"(?:place of business|residing|mailing address) at\]?\s*$", before, re.I):
            return f"{role} Address"

    # 4. Specific local-window patterns (checked before generic words!)
    if blank.startswith("$") and re.search(r"^\s*per year", after, re.I):
        return "Annual Base Salary"
    if blank.startswith("$") and re.search(r"^\s*per month", after, re.I):
        return "Monthly Retainer Fee"
    if blank.startswith("$") and re.search(r"^\s*per\s*\[?week\]?\s*\[?month\]?", after, re.I):
        return "Periodic Service Fee"
    if blank.startswith("$") and re.search(r"^\s*per hour", after, re.I):
        return "Hourly Rate"
    if re.search(r"^\s*hours per week", after, re.I):
        return "Expected Weekly Hours"
    if re.search(r"^\s*days of employment are an introductory", after, re.I):
        return "Introductory Period Days"
    if re.search(r"^\s*days of paid vacation", after, re.I):
        return "Paid Vacation Days Per Year"
    if re.search(r"^\s*weeks[’']?\s+written notice", after, re.I):
        return "Notice Period Weeks"
    if re.search(r"^\s*days[’']?\s+written notice", after, re.I):
        return "Termination Notice Days"
    if re.search(r"^\s*months after", after, re.I):
        return "Non-Solicitation Months"
    if re.search(r"^\s*years after", after, re.I):
        return "Confidentiality Survival Years"

    local_rules: list[tuple[re.Pattern[str], str]] = [
        (re.compile(r"(?:entered into|made) as of\s+$", re.I), "Effective Date"),
        (re.compile(r"Effective from\s+$", re.I), "Commencement Date"),
        (re.compile(r"Party A:\s*$", re.I), "Party A Name"),
        (re.compile(r"Party B:\s*$", re.I), "Party B Name"),
        (re.compile(r"^of\s+$", re.I), "Party Address"),
        (re.compile(r"Party A expects to contribute:\s*$", re.I), "Party A Contribution"),
        (re.compile(r"Party B expects to contribute:\s*$", re.I), "Party B Contribution"),
        (re.compile(r"Indicative timeline and milestones:\s*$", re.I), "Indicative Timeline and Milestones"),
        (re.compile(r"Other key expected terms:\s*$", re.I), "Other Key Expected Terms"),
        (re.compile(r"aim to do so by\s+$", re.I), "Target Definitive Agreement Date"),
        (re.compile(r"^For\s+$", re.I), "Exclusivity Period Days"),
        (re.compile(r"or\s*\(ii\)\s*$", re.I), "Long-Stop Termination Date"),
        (re.compile(r"Start:\s*$", re.I), "Start Date"),
        (re.compile(r"Completion:\s*$", re.I), "Completion Date"),
        (re.compile(r"reporting to\s+$", re.I), "Reporting Manager"),
        (re.compile(r"performing the duties listed in\s+$", re.I), "Duties Schedule Reference"),
        (re.compile(r"workplace is\s+$", re.I), "Workplace Location"),
        (re.compile(r"remote from\s+$", re.I), "Remote Work Location"),
        (re.compile(r"Wages are paid\s+$", re.I), "Pay Frequency"),
        (re.compile(r"bonus as follows:\s*$", re.I), "Bonus Terms"),
        (re.compile(r"before the Effective Date:\s*$", re.I), "Pre-Employment Conditions"),
        (re.compile(r"number of sheets:\s*$", re.I), "Attached Sheets Count"),
        (re.compile(r"Employee[’']?s name:\s*$", re.I), "Employee Name"),
        (re.compile(r"Date:\s*$", re.I), "Date Signed"),
        (re.compile(r"fixed fee of\s+\$?$", re.I), "Fixed Fee Amount"),
        (re.compile(r"fixed project fee[\s\S]*?^\$?$", re.I), "Fixed Fee Amount"),
        (re.compile(r"(?:as follows|payable as follows)\s*:\s*$", re.I), "Payment Schedule"),
        (re.compile(r"estimated total\s*:\s*\$?$", re.I), "Estimated T&M Total"),
        (re.compile(r"not exceed\s+\$?$", re.I), "Fee Cap"),
        (re.compile(r"up to a total of\s+\$?$", re.I), "Expense Reimbursement Cap"),
        (re.compile(r"\$[^\s]+\s*:\s*$", re.I), "Approved Expense Categories"),
        (re.compile(r"reach\s+$", re.I), "Cap Notification Percentage"),
        (re.compile(r"monthly retainer of\s+\$?$", re.I), "Monthly Retainer Fee"),
        (re.compile(r"(?:for up to|covering up to|approximately)\s+$", re.I), "Retainer Hours Per Month"),
        (re.compile(r"(?:and cost|by the company:|hourly rate(?: of)?)\s*\$?$", re.I), "Hourly Rate"),
        (re.compile(r"approved in advance for this sow\s*:\s*$", re.I), "Approved Expenses"),
        (re.compile(r"single expense over\s+\$?$", re.I), "Expense Approval Threshold"),
        (re.compile(r"other terms\s*:\s*$", re.I), "Expense Policy Terms"),
        (re.compile(r"other\s*:\s*$", re.I), "Other Fee or Invoicing Terms"),
        (re.compile(r"by day\s+$", re.I), "Invoice Day of Month"),
        (re.compile(r"due within\s+$", re.I), "Invoice Payment Days"),
        (re.compile(r"pay each invoice within\s+$", re.I), "Invoice Payment Days"),
        (re.compile(r"until paid at\s+$", re.I), "Late Payment Interest Rate"),
        (re.compile(r"re-deliver the deliverable within\s+$", re.I), "Correction Business Days"),
        (re.compile(r"notify the (?:Contractor|Provider) in writing,? within\s+$", re.I), "Acceptance or Defect Notice Days"),
        (re.compile(r"notifies the (?:Contractor|Provider) in writing of a breach[^_]{0,40}within\s+$", re.I), "Warranty Notice Days"),
        (re.compile(r"cure within\s+$", re.I), "Breach Cure Days"),
        (re.compile(r"within\s+$", re.I), "Acceptance Review Business Days"),
        (re.compile(r"or until\s+$", re.I), "End Date"),
        (re.compile(r'write\s+"none"\s+if\s+there\s+are\s+none\s*:\s*$', re.I), "Special Terms"),
        (re.compile(r"laws of the state of\s+$", re.I), "Governing Law State"),
        (re.compile(r"governed by the laws of\s+$", re.I), "Governing Law Jurisdiction"),
        (re.compile(r"courts (?:located )?in\s+$", re.I), "Venue Location"),
    ]
    for pattern, label in local_rules:
        if pattern.search(before):
            return label

    # 5. Generic inline `Label: ____` lookbehind
    colon_match = re.search(r"([A-Za-z][A-Za-z0-9 ’'/-]{2,36})\s*:\s*\$?$", before)
    if colon_match:
        candidate = colon_match.group(1).strip()
        if candidate.lower() not in {"follows", "as follows", "example", "check one", "note"}:
            return candidate.title()

    # 6. Check after-window units
    if re.search(r"%\s*\(80%\s*if left blank\)", after, re.I):
        return "Cap Notification Percentage"
    if re.search(r"%\s*per month", after, re.I):
        return "Late Payment Interest Rate"
    if re.search(r"hours per month", after, re.I):
        return "Retainer Hours Per Month"
    if re.search(r"hours without the Client", after, re.I):
        return "Maximum Authorized Hours"
    if re.search(r"per hour", after, re.I):
        return "Hourly Rate"
    if re.search(r"per month", after, re.I) and blank.startswith("$"):
        return "Monthly Retainer Fee"
    if re.search(r"business days", after, re.I):
        return "Review Business Days"
    if re.search(r"days of receipt|days after receiving", after, re.I):
        return "Invoice Payment Days"

    if blank.startswith("$"):
        return "Fee Amount"
    return "Additional Term"


def _has_underline(run) -> bool:
    try:
        if run.font.underline:
            return True
    except Exception:
        pass
    rpr = run._element.find(qn("w:rPr"))
    if rpr is None:
        return False
    return rpr.find(qn("w:u")) is not None


def _is_form_blank_run(run) -> bool:
    """Underlined spaces/tabs, underscore chars, or empty fill slots."""
    text = run.text or ""
    if re.fullmatch(r"[_\uFF3F]+", text or ""):
        return True
    if not text:
        return False
    if text.strip(" \t") != "":
        return False
    return _has_underline(run) or ("\t" in text)


def _rebuild_paragraph_with_highlights(
    paragraph: Paragraph,
    replacements: list[tuple[int, int, str]],
) -> int:
    """Replace character spans in a paragraph with yellow-highlighted placeholder text."""
    text = paragraph.text or ""
    if not replacements:
        return 0
    replacements = sorted(replacements, key=lambda r: r[0])
    cleaned: list[tuple[int, int, str]] = []
    last_end = -1
    for start, end, repl in replacements:
        if start < last_end or start < 0 or end > len(text) or start >= end:
            continue
        cleaned.append((start, end, repl))
        last_end = end
    if not cleaned:
        return 0

    p_el = paragraph._p
    for child in list(p_el):
        if child.tag == qn("w:r"):
            p_el.remove(child)

    cursor = 0
    count = 0
    for start, end, repl in cleaned:
        if start > cursor:
            paragraph.add_run(text[cursor:start])
        run = paragraph.add_run(repl)
        run.font.highlight_color = YELLOW
        count += 1
        cursor = end
    if cursor < len(text):
        paragraph.add_run(text[cursor:])
    return count


def disambiguate_signature_placeholders(path: Path) -> int:
    """Disambiguate generic [Signatory] / [Title] / [Name] / [Date] / [Fill In] in signature blocks by party role."""
    document = Document(str(path))
    updated = 0
    current_party: str | None = None
    in_change_order_exhibit = False

    for paragraph in _iter_paragraphs(document):
        text = (paragraph.text or "").strip()
        if not text:
            continue
        if _is_change_order_exhibit(text):
            in_change_order_exhibit = True
        if in_change_order_exhibit:
            continue

        detected = _detect_party_header(text)
        if detected:
            current_party = detected

        if not current_party:
            continue

        # Handle `By: [Fill In]` in precedent signature blocks
        if re.match(r"^By\s*:\s*\[Fill In(?:\s*\d+)?\]$", text, re.I):
            for run in paragraph.runs:
                if _is_yellow_run(run) and "fill in" in (run.text or "").lower():
                    run.text = f"[{current_party} Signatory Name]"
                    updated += 1
            continue

        sig_label = _signature_field_label(text, current_party)
        if not sig_label:
            continue

        for run in paragraph.runs:
            if _is_yellow_run(run):
                raw = (run.text or "").strip()
                if raw in {"[Signatory]", "[Name]", "[Title]", "[Date]"} or raw.lower().startswith("[fill in"):
                    run.text = f"[{sig_label}]"
                    updated += 1

    if updated:
        document.save(str(path))
        logger.info("disambiguated {} signature block placeholders in {}", updated, path.name)
    return updated


def highlight_underlined_form_blanks(path: Path) -> int:
    """Highlight Word form blanks made of underlined spaces/tabs (not underscore glyphs)."""
    document = Document(str(path))
    highlighted = 0
    used_labels: dict[str, int] = {}
    prev_text = ""
    current_party: str | None = None
    in_change_order_exhibit = False

    for paragraph in _iter_paragraphs(document):
        runs = list(paragraph.runs)
        text = paragraph.text or ""
        stripped = text.strip()
        if _is_change_order_exhibit(stripped):
            in_change_order_exhibit = True
        if in_change_order_exhibit:
            continue

        detected_party = _detect_party_header(stripped)
        if detected_party:
            current_party = detected_party

        if _is_signature_line(stripped):
            prev_text = stripped
            continue

        if any("_" in (r.text or "") for r in runs):
            if stripped:
                prev_text = stripped
            continue

        if not runs:
            continue

        offsets: list[tuple[int, int]] = []
        cursor = 0
        for run in runs:
            piece = run.text or ""
            offsets.append((cursor, cursor + len(piece)))
            cursor += len(piece)

        groups: list[tuple[int, int]] = []
        i = 0
        while i < len(runs):
            if _is_form_blank_run(runs[i]):
                start = i
                while i < len(runs) and _is_form_blank_run(runs[i]):
                    i += 1
                if start > 0 and (runs[start - 1].text or "").strip() == "$":
                    start -= 1
                groups.append((start, i))
            else:
                i += 1

        spans: list[tuple[int, int, str]] = []
        for start, end in groups:
            char_start = offsets[start][0]
            char_end = offsets[end - 1][1]
            blank_text = text[char_start:char_end]
            left = text[:char_start].strip()
            right = text[char_end:].strip()
            if (
                left.isupper()
                and right.isupper()
                and len(left) <= 20
                and len(right) <= 20
                and "\t" in blank_text
                and "$" not in blank_text
            ):
                continue

            sig_label = _signature_field_label(stripped, current_party)
            if sig_label:
                base = sig_label
            elif char_start <= 1 and right.lower().startswith("client shall"):
                base = "Select Payment Option"
            else:
                base = _label_from_local_context(
                    text,
                    blank_text or "____",
                    char_start,
                    char_end,
                    prev_text,
                )
            used_labels[base] = used_labels.get(base, 0) + 1
            label = base if used_labels[base] == 1 else f"{base} {used_labels[base]}"
            ph = f"[{label}]"
            if char_start > 0 and text[char_start - 1].isalnum():
                ph = " " + ph
            if char_end < len(text) and text[char_end].isalnum():
                ph = ph + " "
            spans.append((char_start, char_end, ph))

        for match in re.finditer(r"\b20xx\b|\bXXXX\b|\bMM/DD/YYYY\b", text, re.I):
            base = "Effective Date"
            used_labels[base] = used_labels.get(base, 0) + 1
            label = base if used_labels[base] == 1 else f"{base} {used_labels[base]}"
            spans.append((match.start(), match.end(), f"[{label}]"))

        highlighted += _rebuild_paragraph_with_highlights(paragraph, spans)
        if stripped:
            prev_text = stripped

    if highlighted:
        document.save(str(path))
        logger.info("highlighted {} underlined form blanks in {}", highlighted, path.name)
    return highlighted


def highlight_underscore_blanks(path: Path) -> int:
    """Turn ____ / $____ / 20xx style blanks into contextual yellow [Label] fields."""
    document = Document(str(path))
    highlighted = 0
    used_labels: dict[str, int] = {}
    prev_text = ""
    prev_had_trailing_blank = False
    current_party: str | None = None
    in_change_order_exhibit = False

    for paragraph in _iter_paragraphs(document):
        text = paragraph.text or ""
        stripped = text.strip()
        if _is_change_order_exhibit(stripped):
            in_change_order_exhibit = True
        if in_change_order_exhibit:
            continue

        detected_party = _detect_party_header(stripped)
        if detected_party:
            current_party = detected_party

        if _is_signature_line(stripped):
            prev_text = stripped
            prev_had_trailing_blank = False
            continue

        # If this paragraph is purely a continuation line of underscores right after a paragraph
        # that already ended with an underscore blank (e.g. Section 10 Special Terms or MOU contributions), clear it.
        if re.fullmatch(r"_+", stripped) and prev_had_trailing_blank:
            for child in list(paragraph._p):
                if child.tag == qn("w:r"):
                    paragraph._p.remove(child)
            prev_had_trailing_blank = False
            continue

        matches = list(BLANK_FIND.finditer(text))
        if not matches:
            if stripped:
                prev_text = stripped
                prev_had_trailing_blank = False
            continue

        yellow_runs = sum(
            1
            for r in paragraph.runs
            if getattr(r.font, "highlight_color", None) == WD_COLOR_INDEX.YELLOW
        )
        if yellow_runs >= len(matches):
            if stripped:
                prev_text = stripped
            prev_had_trailing_blank = False
            continue

        spans: list[tuple[int, int, str]] = []
        sig_label = _signature_field_label(stripped, current_party)
        for match in matches:
            blank = match.group(0)
            if sig_label and len(matches) == 1:
                base = sig_label
            else:
                base = _label_from_local_context(
                    text,
                    blank,
                    match.start(),
                    match.end(),
                    prev_text,
                )
            used_labels[base] = used_labels.get(base, 0) + 1
            label = base if used_labels[base] == 1 else f"{base} {used_labels[base]}"
            placeholder = f"[{label}]"
            spans.append((match.start(), match.end(), placeholder))

        trailing_blank = bool(re.search(r"_{3,}\s*$", stripped)) and not re.fullmatch(r"_+", stripped)
        highlighted += _rebuild_paragraph_with_highlights(paragraph, spans)
        if stripped:
            prev_text = stripped
            prev_had_trailing_blank = trailing_blank

    if highlighted:
        document.save(str(path))
        logger.info("highlighted {} underscore/blank fields in {}", highlighted, path.name)
    return highlighted


def apply_find_replace_highlights(path: Path, fields: list[dict[str, Any]]) -> int:
    """Apply paragraph-scoped AI find→placeholder replacements with yellow highlight."""
    if not fields:
        return 0
    document = Document(str(path))
    highlighted = 0

    by_paragraph: dict[int, list[dict[str, Any]]] = {}
    for field in fields:
        p_idx = field.get("p")
        if isinstance(p_idx, int) and p_idx >= 0:
            by_paragraph.setdefault(p_idx, []).append(field)

    for p_idx, paragraph in enumerate(_iter_paragraphs(document)):
        candidates = by_paragraph.get(p_idx)
        if not candidates:
            continue
        text = paragraph.text or ""
        if not text or _is_signature_line(text):
            continue

        # Fast path: if the target `find` is already an existing yellow run (e.g. `[Additional Term 2]`),
        # rename that yellow run directly in-place without rebuilding the whole paragraph!
        unmatched: list[dict[str, Any]] = []
        for field in candidates:
            find = (field.get("find") or "").strip()
            placeholder = (field.get("placeholder") or "").strip()
            if not find or not placeholder or find == placeholder:
                continue
            if not is_valid_placeholder_token(placeholder):
                continue
            renamed_in_run = False
            for run in paragraph.runs:
                if _is_yellow_run(run) and (run.text or "").strip() == find:
                    run.text = placeholder
                    highlighted += 1
                    renamed_in_run = True
                    break
            if not renamed_in_run:
                unmatched.append(field)

        if not unmatched:
            continue

        # Rebuild paragraph while preserving any EXISTING yellow runs AND adding new yellow spans
        existing_yellow_spans: list[tuple[int, int, str]] = []
        cursor = 0
        for run in paragraph.runs:
            r_text = run.text or ""
            r_len = len(r_text)
            if r_len and _is_yellow_run(run):
                existing_yellow_spans.append((cursor, cursor + r_len, r_text))
            cursor += r_len

        new_spans: list[tuple[int, int, str]] = list(existing_yellow_spans)
        for field in unmatched:
            find = (field.get("find") or "").strip()
            placeholder = (field.get("placeholder") or "").strip()
            start = text.find(find)
            if start < 0:
                continue
            end = start + len(find)
            if any(not (end <= s or start >= e) for s, e, _ in new_spans):
                continue
            new_spans.append((start, end, placeholder))

        if len(new_spans) > len(existing_yellow_spans):
            _rebuild_paragraph_with_highlights(paragraph, new_spans)
            highlighted += len(new_spans) - len(existing_yellow_spans)

    if highlighted:
        document.save(str(path))
        logger.info("applied {} paragraph-scoped AI field highlights in {}", highlighted, path.name)
    return highlighted


def _numbered_paragraphs(path: Path) -> tuple[list[tuple[int, str]], list[tuple[int, str]]]:
    """Return (all_numbered_paragraphs, paragraphs_with_generic_or_raw_blanks)."""
    document = Document(str(path))
    all_paras: list[tuple[int, str]] = []
    generic_paras: list[tuple[int, str]] = []
    for i, paragraph in enumerate(_iter_paragraphs(document)):
        text = (paragraph.text or "").strip()
        if _is_change_order_exhibit(text):
            break
        if not text:
            continue
        all_paras.append((i, text))
        if (
            re.search(r"\[(?:Additional Term|Additional Details|Fill In)(?:\s+\d+)?\]", text, re.I)
            or ("___" in text and not _is_signature_line(text))
        ):
            generic_paras.append((i, text))
    return all_paras, generic_paras


async def ai_refine_and_discover_fields(path: Path) -> int:
    """Use Capco Wayne to:
    1) Rename any generic `[Additional Term]` / `[Fill In]` or raw `___` blanks to semantic `[Labels]`.
    2) If the uploaded template is a precedent contract with < 6 fields, identify hardcoded deal variables
       (Effective Date, Party Names, Addresses, Fees, Governing Law, Signatories) and turn them into yellow `[Fields]`.
    """
    if not copilot.settings.ai_configured:
        return 0

    all_paras, generic_paras = _numbered_paragraphs(path)
    scanned = scan_docx(path)
    applied_total = 0

    # Pass A: Refine generic `[Additional Term ...]` / `[Fill In ...]` or remaining `___`
    if generic_paras:
        lines: list[str] = []
        para_map = dict(all_paras)
        for p_idx, text in generic_paras[:45]:
            prev_line = para_map.get(p_idx - 1, "")
            if prev_line and prev_line != text:
                lines.append(f"[P{p_idx - 1} context] {prev_line[:180]}")
            lines.append(f"[P{p_idx}] {text}")

        system = (
            "You are a legal template engineering AI. Each `[P<num>]` line below is a paragraph from a contract template "
            "that contains generic placeholders like `[Additional Term]`, `[Additional Term 2]`, `[Fill In]`, or raw blanks `____`. "
            "Rename every generic placeholder or raw blank into a specific, concise, unique semantic `[Label]` that accurately "
            "reflects what the lawyer/user must fill in at that exact spot.\n"
            "Return ONLY valid JSON with this shape:\n"
            '{"fields":[{"p":14,"find":"[Additional Term 2]","placeholder":"[Position Title]"}]}\n'
            "Rules:\n"
            "- `p` MUST be the exact integer paragraph number from `[P<num>]`.\n"
            "- `find` MUST be the exact token in `[P<num>]` (e.g. `\"[Additional Term 2]\"` or `\"____\"`).\n"
            "- `placeholder` MUST be `[Title Case Semantic Label]` (2 to 5 words, never `[Additional Term]` or `[Fill In]`).\n"
            "- If a paragraph has multiple placeholders, give each one a distinct semantic name."
        )
        try:
            raw = await copilot._chat(
                system=system,
                context="Mode: template_field_refinement",
                messages=[{"role": "user", "content": "\n".join(lines)}],
            )
            payload = _extract_json(raw)
            fields = payload.get("fields") if isinstance(payload, dict) else None
            if isinstance(fields, list) and fields:
                applied_total += apply_find_replace_highlights(path, fields)
        except Exception as exc:  # noqa: BLE001
            logger.warning("AI generic field refinement failed: {}", exc)

    # Pass B: If the template still has very few fields (< 6), it may be a precedent contract
    # (like CapcoHaseMSA1.docx) where dates, party names, fees, or governing law are hardcoded in plain text.
    scanned = scan_docx(path)
    if len(scanned.fields) < 6 and all_paras:
        doc_dump = "\n".join(f"[P{idx}] {txt}" for idx, txt in all_paras[:80])
        system = (
            "You are a legal template engineering AI. The user uploaded a contract precedent to turn it into a reusable template. "
            "Identify the specific deal variables currently hardcoded in the text (such as the specific Effective Date, "
            "Party A / Provider legal name, Party B / Customer legal name, specific fee amounts, review/payment days, "
            "governing law jurisdiction, and signatory names/titles) that should be turned into yellow `[Placeholders]`.\n"
            "Return ONLY valid JSON with this shape:\n"
            '{"fields":[{"p":1,"find":"October 12, 2026","placeholder":"[Effective Date]"},'
            '{"p":2,"find":"Capco Hong Kong","placeholder":"[Provider Name]"}]}\n'
            "Rules:\n"
            "- `p` MUST be the exact integer paragraph index from `[P<num>]`.\n"
            "- `find` MUST be the exact hardcoded substring inside `[P<num>]` to replace.\n"
            "- Do NOT replace standard defined terms in quotes like `(\"Provider\")`, `(\"Customer\")`, `(\"Effective Date\")` — "
            "replace the actual party name or date value next to them.\n"
            "- NEVER replace `Signature: ____` or `By: ____` lines.\n"
            "- `placeholder` MUST be a clear `[Semantic Label]` in brackets."
        )
        try:
            raw = await copilot._chat(
                system=system,
                context="Mode: precedent_template_discovery",
                messages=[{"role": "user", "content": doc_dump[:18000]}],
            )
            payload = _extract_json(raw)
            fields = payload.get("fields") if isinstance(payload, dict) else None
            if isinstance(fields, list) and fields:
                applied_total += apply_find_replace_highlights(path, fields)
        except Exception as exc:  # noqa: BLE001
            logger.warning("AI precedent field discovery failed: {}", exc)

    return applied_total


def _extract_json(raw: str) -> dict[str, Any]:
    text = raw.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
    try:
        data = json.loads(text)
        return data if isinstance(data, dict) else {"fields": data}
    except json.JSONDecodeError:
        match = re.search(r"\{[\s\S]*\}", text)
        if not match:
            return {}
        try:
            data = json.loads(match.group(0))
            return data if isinstance(data, dict) else {}
        except json.JSONDecodeError:
            return {}


async def prepare_template_fields(path: Path, *, use_ai: bool = True) -> dict[str, Any]:
    """Full pipeline: tables → brackets → signature disambiguation → underscores → form blanks → AI semantic discovery."""
    table_highlights = highlight_table_blanks(path)
    bracketed = ensure_placeholder_highlights(path)
    sig_disambiguated = disambiguate_signature_placeholders(path)
    underscore = highlight_underscore_blanks(path)
    form_blanks = highlight_underlined_form_blanks(path)

    ai_applied = 0
    if use_ai:
        ai_applied = await ai_refine_and_discover_fields(path)

    from app.services.docx_package import enrich_docx_for_editor

    enrich_docx_for_editor(path)

    scanned = scan_docx(path)
    return {
        "table_highlights": table_highlights,
        "bracket_highlights": bracketed,
        "signature_disambiguated": sig_disambiguated,
        "form_blank_highlights": form_blanks,
        "underscore_highlights": underscore,
        "ai_highlights": ai_applied,
        "field_count": len(scanned.fields),
        "fields": list(scanned.fields.values()),
    }
