"""Uploaded document store for review / assessment / completion workflow."""

from __future__ import annotations

import json
import re
import uuid
from datetime import datetime, timezone
from pathlib import Path

from fastapi import UploadFile

from app.config import Settings, get_settings
from app.services.docx_package import enrich_docx_for_editor
from app.services.editor_content import docx_to_blocks
from app.services.field_discovery import prepare_template_fields


def _display_title(filename: str | None, fallback: str) -> str:
    raw = (filename or "").strip() or fallback
    stem = Path(raw).stem.strip() or fallback
    cleaned = re.sub(r"[\s_]+", " ", stem).strip()
    return cleaned or fallback


class DocumentStore:
    def __init__(self, settings: Settings | None = None) -> None:
        self.settings = settings or get_settings()
        self.index_path = self.settings.data_dir / "uploads_index.json"
        self._items: dict[str, dict] = {}
        self._load()

    def _load(self) -> None:
        if not self.index_path.exists():
            self._items = {}
            return
        self._items = json.loads(self.index_path.read_text(encoding="utf-8"))

    def _save(self) -> None:
        self.index_path.parent.mkdir(parents=True, exist_ok=True)
        self.index_path.write_text(json.dumps(self._items, indent=2), encoding="utf-8")

    def list_documents(self) -> list[dict]:
        return sorted(self._items.values(), key=lambda row: row["created_at"], reverse=True)

    def get(self, document_id: str) -> dict | None:
        return self._items.get(document_id)

    def refresh_from_disk(self, document_id: str) -> dict | None:
        """Re-scan an uploaded document after AI field discovery updates its .docx on disk."""
        row = self._items.get(document_id)
        if not row:
            return None
        path = Path(row["path"])
        if not path.exists():
            return row
        editor = docx_to_blocks(path)
        row["updated_at"] = datetime.now(timezone.utc).isoformat()
        row["plain_text"] = editor.get("plain_text") or ""
        row["blocks"] = editor.get("blocks") or []
        row["fields"] = editor.get("fields") or []
        row["field_count"] = len(row["fields"])
        self._save()
        return row

    async def save_upload(self, upload: UploadFile, *, mode: str = "review") -> dict:
        """
        Persist an uploaded .docx for workspace open.

        - Assessment uploads (`mode == "assess"` or `"extract"`) keep the ORIGINAL bytes unchanged.
        - Completion / draft / template-style uploads (`mode in {"complete", "draft", "fill", "sign"}`)
          run `prepare_template_fields(..., use_ai=True)` so all blanks, tables, and variables are
          automatically identified and yellow-highlighted for Capco Wayne.
        """
        original_name = upload.filename or "document.docx"
        suffix = Path(original_name).suffix.lower() or ".docx"
        if suffix not in {".docx"}:
            raise ValueError("Only .docx uploads are supported in this release")

        document_id = f"doc_{uuid.uuid4().hex[:12]}"
        safe_base = re.sub(r"[^\w.\- ]+", "_", Path(original_name).name).strip() or f"{document_id}.docx"
        if not safe_base.lower().endswith(".docx"):
            safe_base = f"{safe_base}.docx"
        dest = self.settings.uploads_dir / document_id / safe_base
        dest.parent.mkdir(parents=True, exist_ok=True)
        content = await upload.read()
        dest.write_bytes(content)

        assess_only = mode in {"assess", "extract"}
        discovery = None
        if not assess_only:
            try:
                discovery = await prepare_template_fields(dest, use_ai=True)
            except Exception:  # noqa: BLE001
                pass
        else:
            try:
                enrich_docx_for_editor(dest)
            except Exception:  # noqa: BLE001
                pass

        editor = docx_to_blocks(dest)
        title = _display_title(original_name, document_id)
        if title.lower() in {"document", "upload", "file", "untitled"}:
            first_line = (editor.get("plain_text") or "").strip().splitlines()[:1]
            if first_line and len(first_line[0].strip()) > 3:
                title = first_line[0].strip()[:120]

        now = datetime.now(timezone.utc).isoformat()
        row = {
            "id": document_id,
            "filename": original_name,
            "title": title,
            "mode": mode,
            "path": str(dest.resolve()),
            "created_at": now,
            "updated_at": now,
            "field_count": 0 if assess_only else len(editor.get("fields") or []),
            "plain_text": editor.get("plain_text") or "",
            "blocks": editor.get("blocks") or [],
            "fields": [] if assess_only else (editor.get("fields") or []),
            "discovery": discovery,
            "status": "uploaded_for_assessment" if assess_only else "uploaded_for_completion",
            "preserve_original": assess_only,
        }
        self._items[document_id] = row
        self._save()
        return row


document_store = DocumentStore()
