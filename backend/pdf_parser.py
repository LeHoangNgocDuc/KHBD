from __future__ import annotations

from pathlib import Path


def inspect_pdf(path: Path) -> dict:
    import fitz
    pages = []
    headings = []
    with fitz.open(path) as doc:
        for i, page in enumerate(doc):
            text = page.get_text("text")
            images = page.get_images(full=True)
            page_headings = [line.strip() for line in text.splitlines()
                             if line.strip() and (line.strip().upper().startswith(("CHƯƠNG", "BÀI ", "HOẠT ĐỘNG", "LUYỆN TẬP", "VẬN DỤNG")))]
            headings.extend({"page": i + 1, "text": h} for h in page_headings)
            pages.append({"page": i + 1, "text": text, "image_count": len(images), "is_textless": not bool(text.strip())})
    return {"filename": path.name, "page_count": len(pages), "pages": pages, "outline_candidates": headings}


def extract_page_images(path: Path, page_number: int, output_dir: Path) -> list[dict]:
    import fitz
    output_dir.mkdir(parents=True, exist_ok=True)
    out = []
    with fitz.open(path) as doc:
        if page_number < 1 or page_number > len(doc):
            raise ValueError("Số trang ngoài phạm vi tài liệu.")
        page = doc[page_number - 1]
        for n, info in enumerate(page.get_images(full=True), start=1):
            xref = info[0]
            pix = fitz.Pixmap(doc, xref)
            if pix.n - pix.alpha > 3:
                pix = fitz.Pixmap(fitz.csRGB, pix)
            name = f"page-{page_number}-image-{n}.png"
            dest = output_dir / name
            pix.save(dest)
            out.append({"filename": name, "path": str(dest)})
    return out


def fallback_note(description: str, page_number: int) -> str:
    return f"[HÌNH VẼ: {description.strip() or 'Bổ sung mô tả hình minh họa'} - Trang {page_number} trong SGK]"
