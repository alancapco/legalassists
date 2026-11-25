"""Template catalog API."""

import base64
from pathlib import Path

from fastapi import APIRouter, File, Form, HTTPException, UploadFile
from pydantic import BaseModel, Field

from app.models.schemas import TemplateDetail, TemplateSummary
from app.services.catalog import STANDARD_CATEGORIES, catalog
from app.services.documents import document_store

router = APIRouter(prefix="/templates", tags=["templates"])


class TemplateUpdateRequest(BaseModel):
    name: str | None = None
    category: str | None = None
    agreement_type: str | None = None
    jurisdiction: str | None = None
    entity_scope: str | None = None
    description: str | None = None


class SaveFromWorkspaceRequest(BaseModel):
    name: str
    category: str = "Other"
    docx_base64: str | None = None
    template_id: str | None = None
    document_id: str | None = None
    use_ai: bool = True


@router.get("", response_model=list[TemplateSummary])
async def list_templates() -> list[TemplateSummary]:
    return catalog.list_templates()


@router.get("/categories")
async def list_categories() -> list[dict]:
    """Contract categories (NDA, MSA, …) with counts for the browse modal."""
    return catalog.list_categories()


@router.get("/meta/standard-categories")
async def standard_categories() -> list[str]:
    return list(STANDARD_CATEGORIES)


@router.post("/reindex", response_model=list[TemplateSummary])
async def reindex_templates() -> list[TemplateSummary]:
    return catalog.refresh()


@router.post("/upload", response_model=TemplateDetail)
async def upload_template(
    file: UploadFile = File(...),
    category: str = Form("Other"),
    name: str | None = Form(None),
) -> TemplateDetail:
    """Upload a .docx as a template in the chosen category (NDA, MSA, …)."""
    filename = file.filename or "template.docx"
    suffix = filename.lower().rsplit(".", 1)[-1] if "." in filename else ""
    if suffix not in {"docx"}:
        raise HTTPException(status_code=400, detail="Only .docx files are supported")
    content = await file.read()
    if not content:
        raise HTTPException(status_code=400, detail="Empty file")
    if len(content) > 5 * 1024 * 1024:
        raise HTTPException(status_code=400, detail="Max size is 5.00 MB")
    try:
        return await catalog.save_upload(
            filename=filename,
            content=content,
            category=category,
            name=name,
            use_ai=True,
        )
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@router.post("/from-workspace", response_model=TemplateDetail)
async def save_template_from_workspace(body: SaveFromWorkspaceRequest) -> TemplateDetail:
    """Save the current workspace .docx as a reusable template with AI field discovery."""
    content: bytes | None = None
    if body.docx_base64:
        try:
            content = base64.b64decode(body.docx_base64)
        except Exception as exc:  # noqa: BLE001
            raise HTTPException(status_code=400, detail="Invalid docx_base64") from exc
    elif body.document_id:
        doc = document_store.get(body.document_id)
        if doc and Path(doc["path"]).exists():
            content = Path(doc["path"]).read_bytes()
    elif body.template_id:
        path = catalog.resolve_path(body.template_id)
        if path and path.exists():
            content = path.read_bytes()

    if not content:
        raise HTTPException(status_code=400, detail="No document content available to save as template")

    safe_filename = f"{body.name.strip() or 'template'}.docx"
    try:
        return await catalog.save_upload(
            filename=safe_filename,
            content=content,
            category=body.category,
            name=body.name.strip() or None,
            use_ai=body.use_ai,
        )
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@router.patch("/{template_id}", response_model=TemplateDetail)
async def update_template(template_id: str, body: TemplateUpdateRequest) -> TemplateDetail:
    detail = catalog.update_template(
        template_id,
        name=body.name,
        category=body.category,
        agreement_type=body.agreement_type,
        jurisdiction=body.jurisdiction,
        entity_scope=body.entity_scope,
        description=body.description,
    )
    if not detail:
        raise HTTPException(status_code=404, detail="Template not found")
    return detail


@router.delete("/{template_id}")
async def delete_template(template_id: str) -> dict:
    if not catalog.delete_template(template_id):
        raise HTTPException(status_code=404, detail="Template not found")
    return {"ok": True, "id": template_id}


@router.get("/{template_id}", response_model=TemplateDetail)
async def get_template(template_id: str) -> TemplateDetail:
    detail = catalog.get(template_id)
    if not detail:
        raise HTTPException(status_code=404, detail="Template not found")
    return detail
