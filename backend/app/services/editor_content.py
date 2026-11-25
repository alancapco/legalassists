"""Convert Word templates/uploads into editor blocks for the workspace canvas."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml.ns import qn

from app.services.yellow_fields import _iter_paragraphs, derive_field_key


def _is_yellow_run(run) -> bool:
    try:
        if run.font.highlight_color == WD_COLOR_INDEX.YELLOW:
            return True
    except Exception:
        pass
    rpr = run._element.find(qn("w:rPr"))
    if rpr is None:
        return False
    highlight = rpr.find(qn("w:highlight"))
    if highlight is None:
        return False
    return (highlight.get(qn("w:val")) or "").lower() == "yellow"


def docx_to_blocks(path: Path) -> dict[str, Any]:
    document = Document(str(path))
    blocks: list[dict[str, Any]] = []
    fields: dict[str, dict[str, Any]] = {}
    prev_non_empty = ""

    for p_index, paragraph in enumerate(_iter_paragraphs(document)):
        style_name = paragraph.style.name if paragraph.style else "Normal"
        level = 0
        if style_name.startswith("Heading"):
            try:
                level = int(style_name.replace("Heading", "").strip() or "1")
            except ValueError:
                level = 1

        para_text = paragraph.text.strip()
        runs_out: list[dict[str, Any]] = []
        runs = paragraph.runs
        i = 0
        while i < len(runs):
            if _is_yellow_run(runs[i]):
                chunks: list[str] = []
                while i < len(runs) and _is_yellow_run(runs[i]):
                    chunks.append(runs[i].text or "")
                    i += 1
                placeholder = "".join(chunks).strip()
                if not placeholder:
                    continue
                key, label = derive_field_key(placeholder)
                if prev_non_empty and (para_text == placeholder or len(para_text) <= len(placeholder) + 12):
                    sample_context = f"{prev_non_empty} → {para_text}"[:240]
                else:
                    sample_context = para_text[:240]
                fields.setdefault(
                    key,
                    {
                        "key": key,
                        "label": label,
                        "placeholder": placeholder,
                        "value": "",
                        "sample_context": sample_context or None,
                    },
                )
                runs_out.append(
                    {
                        "type": "field",
                        "key": key,
                        "label": label,
                        "placeholder": placeholder,
                        "text": placeholder,
                    }
                )
                continue

            text = runs[i].text or ""
            if text:
                runs_out.append(
                    {
                        "type": "text",
                        "text": text,
                        "bold": bool(runs[i].bold),
                        "italic": bool(runs[i].italic),
                    }
                )
            i += 1

        if para_text:
            prev_non_empty = para_text

        if not runs_out and not para_text:
            blocks.append({"type": "spacer", "index": p_index})
            continue

        blocks.append(
            {
                "type": "paragraph",
                "index": p_index,
                "heading_level": level,
                "runs": runs_out
                or [{"type": "text", "text": paragraph.text, "bold": False, "italic": False}],
            }
        )

    plain = "\n".join(
        "".join(run.get("text", "") for run in block.get("runs", []))
        for block in blocks
        if block.get("type") == "paragraph"
    )
    return {
        "blocks": blocks,
        "fields": list(fields.values()),
        "plain_text": plain,
        "title": path.stem.replace("_", " ").title(),
    }
