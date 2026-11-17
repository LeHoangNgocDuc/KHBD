from __future__ import annotations

import json
import re
import uuid
from datetime import datetime, timedelta
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from .docx_generator import generate_docx
from .merger import merge_docx
from .models import LessonMeta, MergeRequest
from .pdf_parser import extract_page_images, fallback_note, inspect_pdf
from .ppct_parser import parse_ppct

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
UPLOADS = DATA / "uploads"
GENERATED = DATA / "generated"
for folder in (UPLOADS, GENERATED): folder.mkdir(parents=True, exist_ok=True)

app = FastAPI(title="Smart KHBD Toán 9", version="0.1.0")
assets: dict[str, dict] = {}
outputs: dict[str, Path] = {}

ALLOWED = {".pdf", ".docx", ".xlsx"}
MAX_SIZE = 60 * 1024 * 1024


def safe_name(name: str) -> str:
    return re.sub(r"[^\w.-]+", "_", Path(name).name, flags=re.UNICODE)[:120] or "upload"


@app.get("/api/health")
def health():
    return {"ok": True, "app": "Smart KHBD Toán 9", "version": app.version}


@app.get("/")
def home():
    return FileResponse(BASE.parent / "frontend" / "index.html")


@app.post("/api/upload")
async def upload(kind: str = Form(...), file: UploadFile = File(...)):
    ext = Path(file.filename or "").suffix.lower()
    if ext not in ALLOWED:
        raise HTTPException(400, "Chỉ nhận PDF, DOCX hoặc XLSX.")
    if kind not in {"ppct", "sgk", "template"}:
        raise HTTPException(400, "Loại tệp không hợp lệ.")
    if kind == "sgk" and ext != ".pdf":
        raise HTTPException(400, "SGK hiện chỉ nhận PDF.")
    if kind == "template" and ext != ".docx":
        raise HTTPException(400, "KHBD mẫu cần là DOCX.")
    file_id = uuid.uuid4().hex
    dest = UPLOADS / f"{file_id}_{safe_name(file.filename or 'upload') }"
    total = 0
    with dest.open("wb") as out:
        while chunk := await file.read(1024 * 1024):
            total += len(chunk)
            if total > MAX_SIZE:
                dest.unlink(missing_ok=True)
                raise HTTPException(413, "Tệp vượt quá 60 MB.")
            out.write(chunk)
    record = {"id": file_id, "kind": kind, "name": file.filename, "path": dest}
    assets[file_id] = record
    result = {"id": file_id, "name": file.filename, "kind": kind, "size": total}
    try:
        if kind == "ppct": result["parse"] = parse_ppct(dest)
        elif kind == "sgk":
            index = inspect_pdf(dest)
            assets[file_id]["index"] = index
            result["outline"] = {"filename": index["filename"], "page_count": index["page_count"],
                                 "outline_candidates": index["outline_candidates"],
                                 "textless_pages": [p["page"] for p in index["pages"] if p["is_textless"]]}
    except Exception as exc:
        result["warning"] = f"Đã lưu tệp nhưng chưa đọc được nội dung: {exc}"
    return result


@app.get("/api/sgk/{file_id}/images/{page_number}")
def page_images(file_id: str, page_number: int):
    item = assets.get(file_id)
    if not item or item["kind"] != "sgk":
        raise HTTPException(404, "Không tìm thấy SGK đã nạp.")
    try:
        images = extract_page_images(item["path"], page_number, GENERATED / file_id)
    except Exception as exc:
        raise HTTPException(400, str(exc))
    return {"page": page_number, "images": images,
            "fallback": fallback_note("Mô tả hình cần giáo viên bổ sung nếu trang không có ảnh tách riêng.", page_number) if not images else None}


@app.get("/api/sgk/{file_id}/pages/{page_number}")
def page_text(file_id: str, page_number: int):
    item = assets.get(file_id)
    if not item or item["kind"] != "sgk":
        raise HTTPException(404, "Không tìm thấy SGK đã nạp.")
    index = item.get("index") or inspect_pdf(item["path"])
    if page_number < 1 or page_number > len(index["pages"]):
        raise HTTPException(400, "Số trang ngoài phạm vi tài liệu.")
    page = index["pages"][page_number - 1]
    return {"page": page_number, "text": page["text"], "image_count": page["image_count"]}


@app.post("/api/generate")
def generate(meta: LessonMeta, template_id: str | None = None, sgk_id: str | None = None):
    template = None
    if template_id:
        item = assets.get(template_id)
        if not item or item["kind"] != "template":
            raise HTTPException(404, "Không tìm thấy KHBD mẫu đã tải lên.")
        template = item["path"]
    sgk_path = None
    if sgk_id:
        sgk = assets.get(sgk_id)
        if not sgk or sgk["kind"] != "sgk":
            raise HTTPException(404, "Không tìm thấy SGK đã nạp.")
        sgk_path = sgk["path"]
    if not meta.prepare_date and meta.date_start:
        try:
            meta.prepare_date = (datetime.strptime(meta.date_start, "%d/%m/%Y") - timedelta(days=3)).strftime("%d/%m/%Y")
        except ValueError:
            pass
    ending = meta.lesson_number_end or meta.lesson_number
    period_part = f"T{meta.lesson_number}" if ending == meta.lesson_number else f"T{meta.lesson_number}_{ending}"
    filename = f"{meta.subject_code}{meta.grade}_T{meta.week}_{period_part}_B{meta.lesson_index}.docx"
    output = GENERATED / f"{uuid.uuid4().hex}_{filename}"
    try:
        result = generate_docx(meta, template, output, sgk_path=sgk_path)
    except Exception as exc:
        raise HTTPException(422, f"Không thể tạo DOCX: {exc}")
    file_id = uuid.uuid4().hex
    outputs[file_id] = output
    return {"id": file_id, "filename": filename, "download_url": f"/api/download/{file_id}",
            "converted_equations": result["converted_equations"], "template_based": result["template_based"],
            "message": "Đã tạo giáo án DOCX.",
            "content_note": "Nếu dùng DOCX mẫu, phần nội dung bài được giữ từ mẫu; cần chọn mẫu đúng bài hoặc tự bổ sung nội dung trước khi phát hành."}


@app.get("/api/download/{file_id}")
def download(file_id: str):
    path = outputs.get(file_id)
    if not path or not path.exists():
        raise HTTPException(404, "Không tìm thấy tệp kết quả.")
    return FileResponse(path, filename=path.name.split("_", 1)[-1], media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.get("/api/generated")
def generated_files():
    return [{"id": key, "filename": path.name.split("_", 1)[-1]} for key, path in outputs.items() if path.exists()]


@app.post("/api/merge")
def merge(request: MergeRequest):
    paths = []
    for file_id in request.file_ids:
        path = outputs.get(file_id)
        if not path or not path.exists():
            raise HTTPException(404, f"Không tìm thấy giáo án {file_id}.")
        paths.append(path)
    merged = GENERATED / f"{uuid.uuid4().hex}_KHBD_Toan_Ghep.docx"
    try:
        merge_docx(paths, merged)
    except Exception as exc:
        raise HTTPException(422, f"Không ghép được DOCX: {exc}")
    file_id = uuid.uuid4().hex
    outputs[file_id] = merged
    return {"id": file_id, "filename": "KHBD_Toan_Ghep.docx", "download_url": f"/api/download/{file_id}"}


app.mount("/static", StaticFiles(directory=BASE.parent / "frontend"), name="frontend")
