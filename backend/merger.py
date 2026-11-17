from pathlib import Path


def merge_docx(files: list[Path], output_path: Path) -> Path:
    if not files:
        raise ValueError("Chọn ít nhất một giáo án để ghép.")
    try:
        from docxcompose.composer import Composer
    except ImportError as exc:
        raise RuntimeError("Thiếu docxcompose. Cài các gói trong requirements.txt.") from exc
    from docx import Document
    master = Document(files[0])
    composer = Composer(master)
    for path in files[1:]:
        composer.doc.add_page_break()
        composer.append(Document(path))
    output_path.parent.mkdir(parents=True, exist_ok=True)
    composer.save(str(output_path))
    return output_path
