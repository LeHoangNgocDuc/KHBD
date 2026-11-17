from __future__ import annotations

import re
from pathlib import Path
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .models import LessonMeta
from .omml_converter import convert_document_equations
from .pdf_parser import extract_page_images, fallback_note


def _all_paragraphs(doc):
    yield from doc.paragraphs
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                yield from cell.paragraphs


def _set_text(paragraph, value: str):
    # Keep the paragraph properties and first run's direct character styling.
    run = paragraph.runs[0] if paragraph.runs else None
    style = run.style if run else None
    paragraph.clear()
    new = paragraph.add_run(value)
    if style:
        new.style = style


def _remove_paragraph(paragraph):
    el = paragraph._element
    parent = el.getparent()
    if parent is not None:
        parent.remove(el)


def _insert_before(doc, marker: str, text: str):
    for p in doc.paragraphs:
        if marker.casefold() in p.text.casefold():
            return p.insert_paragraph_before(text)
    return doc.add_paragraph(text)


def _set_header(doc, meta: LessonMeta):
    items = list(doc.paragraphs)
    replacements = [
        ("Ngày soạn:", f"Ngày soạn: {meta.prepare_date or '____/____/________'}"),
        ("Tuần ", f"Tuần {meta.week}."),
        ("CHƯƠNG", f"{meta.chapter} {meta.chapter_title}".strip()),
        ("BÀI ", f"BÀI {meta.lesson_index}. {meta.lesson_title.upper()} ({meta.lesson_periods} TIẾT)"),
        ("Môn:", f"Môn: Toán {meta.grade} ({'Đại số' if meta.subject_code == 'DS' else 'Hình học'});"),
        ("Lớp", f"Lớp : {meta.classes}"),
        ("Thời gian thực hiện:", f"Thời gian thực hiện: {meta.lesson_periods} tiết"),
    ]
    used = set()
    for marker, value in replacements:
        for p in items:
            if id(p) in used:
                continue
            if marker.casefold() in p.text.casefold():
                _set_text(p, value); used.add(id(p)); break
    for p in list(doc.paragraphs):
        if "(Đã soạn ở Tuần" in p.text:
            _remove_paragraph(p)
    if meta.previous_week:
        _insert_before(doc, "Môn:", f"(Đã soạn ở Tuần {meta.previous_week}" +
                       (f" - Tiết {meta.previous_periods}" if meta.previous_periods else "") + ")")


def _prune_optional_blocks(doc, meta: LessonMeta):
    # Remove the full inclusive education objectives section as one block so no
    # orphaned goals remain between the normal objectives and teaching materials.
    if not meta.enable_gdhn:
        in_block = False
        for p in list(doc.paragraphs):
            text = p.text.strip().upper()
            if text.startswith("* MỤC TIÊU DÀNH CHO H"):
                in_block = True
            if in_block and text.startswith("II. THIẾT BỊ"):
                in_block = False
            elif in_block:
                _remove_paragraph(p)
    paragraphs = list(_all_paragraphs(doc))
    # Drop or normalize integration paragraphs across body and activity tables.
    for p in paragraphs:
        text = p.text.strip()
        if not meta.enable_nls and text.startswith("Bước 1:") and "[Tích hợp NLS]" in text:
            _set_text(p, text[:text.index("[Tích hợp NLS]")].rstrip())
            continue
        if not meta.enable_ai and text.startswith("Bước 1:") and "💡 Prompt gợi ý" in text:
            _set_text(p, text[:text.index("💡 Prompt gợi ý")].rstrip())
            text = p.text.strip()
        if not meta.enable_gdhn and "[GDHN]" in text:
            if text.startswith("Bước 2:"):
                _set_text(p, text[:text.index("[GDHN]")].rstrip())
                continue
            _remove_paragraph(p)
            continue
        if not meta.enable_nls and ("[Tích hợp NLS]" in text or "Năng lực số (theo" in text or re.search(r"\[TC\d", text)):
            _remove_paragraph(p)
        elif not meta.enable_ai and ("Năng lực trí tuệ nhân tạo" in text or "[9.D" in text or "[9.B" in text or "Prompt AI" in text or "Prompt gợi ý" in text):
            _remove_paragraph(p)
        elif not meta.enable_gdhn and ("[GDHN" in text or "MỤC TIÊU DÀNH CHO H" in text):
            _remove_paragraph(p)

    if meta.enable_nls and not any("Năng lực số (theo" in p.text for p in _all_paragraphs(doc)):
        _insert_before(doc, "3. Phẩm chất", "* Năng lực số (theo TT 02/2025/TT-BGDĐT & Hướng dẫn QĐ 2422/QĐ-BGDĐT):\n+ Khai thác thông tin [TC2.1.1]; lựa chọn công cụ số giải quyết vấn đề [TC2.5.2]; lưu trữ kết quả [TC2.1.3].")
    if meta.enable_ai and not any("Năng lực trí tuệ nhân tạo (AI)" in p.text for p in _all_paragraphs(doc)):
        _insert_before(doc, "3. Phẩm chất", "* Năng lực trí tuệ nhân tạo (AI) (theo QĐ 2422/QĐ-BGDĐT):\n+ Kiểm chứng đề xuất AI [9.D1.1]; sử dụng AI có trách nhiệm [9.B3.1].")

    if meta.enable_gdhn:
        for p in _all_paragraphs(doc):
            if "MỤC TIÊU DÀNH CHO H" in p.text:
                _set_text(p, "* MỤC TIÊU DÀNH CHO HỌC SINH KHUYẾT TẬT")
            elif "Điều chỉnh mục tiêu Giáo dục hòa nhập" in p.text:
                _set_text(p, f"* Điều chỉnh mục tiêu Giáo dục hòa nhập (Cho HSKT {meta.disability}):")
        if not any("MỤC TIÊU DÀNH CHO H" in p.text for p in doc.paragraphs):
            block = [
                "* MỤC TIÊU DÀNH CHO HỌC SINH KHUYẾT TẬT",
                "1. Kiến thức: Đạt yêu cầu cốt lõi của bài với nhiệm vụ đã điều chỉnh.",
                "2. Năng lực: Tham gia hoạt động phù hợp khả năng, có hỗ trợ.",
                "3. Phẩm chất: Tự tin, hợp tác và tích cực tham gia.",
                f"* Điều chỉnh mục tiêu Giáo dục hòa nhập (HS khuyết tật {meta.disability}):",
                "+ [GDHN.KT]: Nhận biết nội dung trọng tâm bằng học liệu phù hợp và hướng dẫn từng bước.",
                "+ [GDHN.KN]: Tương tác với bạn đồng hành, hoàn thành phần việc phù hợp.",
            ]
            for line in block:
                _insert_before(doc, "II. THIẾT BỊ", line)
        for p in list(_all_paragraphs(doc)):
            if p.text.startswith("Bước 2: Thực hiện nhiệm vụ") and "[GDHN]" not in p.text:
                p.add_run("\n- [GDHN] Biện pháp hỗ trợ HSKT: giao nhiệm vụ theo khả năng, cung cấp hướng dẫn từng bước và bố trí bạn đồng hành; điều chỉnh học liệu cho dạng khuyết tật "+meta.disability+".")

    for p in list(_all_paragraphs(doc)):
        if p.text.startswith("Bước 1:"):
            if meta.enable_nls and "[Tích hợp NLS]" not in p.text:
                p.add_run("\n[Tích hợp NLS] [TC2.5.2] Hướng dẫn HS chọn công cụ số phù hợp để kiểm tra kết quả nhiệm vụ.")
            if meta.enable_ai and "Prompt gợi ý" not in p.text:
                p.add_run("\n💡 Prompt gợi ý cho GV/HS: Hãy giải thích từng bước và nêu cách tự kiểm tra kết quả; không chấp nhận đáp án AI nếu chưa đối chiếu kiến thức đã học.")
    if meta.enable_nls and not any("Nhiệm vụ NLS" in p.text for p in _all_paragraphs(doc)):
        doc.add_paragraph("- Nhiệm vụ NLS: Dùng công cụ số phù hợp để luyện tập, lưu kết quả và ghi nguồn học liệu.")
    if meta.enable_ai and not any("Prompt gợi ý" in p.text for p in _all_paragraphs(doc)):
        doc.add_paragraph("💡 Prompt gợi ý về nhà: Hãy tạo một bài tập tương tự chủ đề hôm nay, trình bày lời giải từng bước và tự kiểm tra kết quả.")


def _header_only(meta: LessonMeta):
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Cm(1.5); sec.bottom_margin = Cm(2.5); sec.left_margin = Cm(2); sec.right_margin = Cm(1.5)
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"; style.font.size = Pt(13)
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    for text, bold in [
        (f"Ngày soạn: {meta.prepare_date}", False), (f"Tuần {meta.week}.", False),
        (f"{meta.chapter} {meta.chapter_title}".strip(), True),
        (f"BÀI {meta.lesson_index}. {meta.lesson_title.upper()} ({meta.lesson_periods} TIẾT)", True),
        (f"(Đã soạn ở Tuần {meta.previous_week} - Tiết {meta.previous_periods})" if meta.previous_week else "", False),
        (f"Môn: Toán {meta.grade} ({'Đại số' if meta.subject_code == 'DS' else 'Hình học'});", False),
        (f"Lớp : {meta.classes}", False), (f"Thời gian thực hiện: {meta.lesson_periods} tiết", False),
    ]:
        if text and text not in ("CHƯƠNG", "BÀI"):
            p = doc.add_paragraph()
            r = p.add_run(text); r.bold = bold
            if bold: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("Nội dung bài học đã được soạn ở tuần trước. Đây là khung tham chiếu PPCT.")
    return doc


def _shade(cell, fill="D9EAF7"):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd"); shd.set(qn("w:fill"), fill); tc_pr.append(shd)


def _set_borders(table):
    tbl_pr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = OxmlElement(f"w:{edge}"); tag.set(qn("w:val"), "single")
        tag.set(qn("w:sz"), "6"); tag.set(qn("w:color"), "404040"); borders.append(tag)
    tbl_pr.append(borders)


def _build_structured(meta: LessonMeta):
    data = meta.lesson_content or {}
    doc = _header_only(meta)
    doc.paragraphs[-1]._element.getparent().remove(doc.paragraphs[-1]._element)
    def heading(text):
        p = doc.add_paragraph(); r=p.add_run(text); r.bold=True
    def line(text):
        if text: doc.add_paragraph(text)
    heading("I. MỤC TIÊU:")
    for title, key in [("1. Kiến thức:", "knowledge"), ("2. Năng lực:", "competencies"), ("3. Phẩm chất:", "qualities")]:
        heading(title)
        values = data.get(key, [])
        if isinstance(values, str): values = [values]
        for value in values: line(value)
    _prune_optional_blocks(doc, meta)
    heading("II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU")
    line("1 - GV: SGK, tài liệu dạy học, thiết bị trình chiếu và học liệu của bài.")
    line("2 - HS: SGK, vở ghi, đồ dùng học tập.")
    heading("III. TIẾN TRÌNH DẠY HỌC")
    activities = data.get("activities", [])
    if not activities:
        activities = [{"title":"A. HOẠT ĐỘNG KHỞI ĐỘNG (MỞ ĐẦU)"},
                       {"title":"B. HÌNH THÀNH KIẾN THỨC MỚI"},
                       {"title":"C. HOẠT ĐỘNG LUYỆN TẬP"},
                       {"title":"D. HOẠT ĐỘNG VẬN DỤNG"}]
    for activity in activities:
        heading(activity.get("title", "Hoạt động học"))
        for key, label in [("objective","a) Mục tiêu"),("content","b) Nội dung"),("product","c) Sản phẩm")]:
            line(f"{label}: {activity.get(key, '')}")
        line("d) Tổ chức thực hiện:")
        table = doc.add_table(rows=1, cols=2); table.alignment=WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False; table.columns[0].width=Cm(8.2); table.columns[1].width=Cm(8.2)
        heads = table.rows[0].cells
        heads[0].text="HĐ CỦA GV VÀ HS"; heads[1].text="SẢN PHẨM DỰ KIẾN"
        for cell in heads:
            _shade(cell)
            for p in cell.paragraphs:
                for r in p.runs: r.bold=True
        row=table.add_row().cells
        steps=activity.get("steps", {})
        for n,label in enumerate(("Chuyển giao nhiệm vụ","Thực hiện nhiệm vụ","Báo cáo, thảo luận","Kết luận, nhận định"),1):
            row[0].add_paragraph(f"Bước {n}: {label}: {steps.get(str(n), 'Giáo viên giao nhiệm vụ; học sinh thực hiện, báo cáo và nhận xét.')}")
        products=activity.get("expected_products", [])
        if isinstance(products,str): products=[products]
        row[1].text="\n".join(products) or "Bổ sung lời giải, sản phẩm học tập từ nội dung bài dạy."
        for cell in row: cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.TOP
        _set_borders(table)
    heading("* HƯỚNG DẪN VỀ NHÀ")
    for task in data.get("homework", ["Ghi nhớ kiến thức trong bài.", "Hoàn thành bài tập trong SBT."]): line("- "+task)
    return doc


def _attach_requested_images(doc, requests: list[dict], sgk_path: Path | None, output_path: Path):
    for request in requests:
        page = int(request.get("page", 0) or 0)
        description = str(request.get("description", "")).strip()
        anchor = str(request.get("anchor", "")).strip()
        target = next((p for p in _all_paragraphs(doc) if anchor and anchor.casefold() in p.text.casefold()), None)
        extracted = []
        if sgk_path and page > 0:
            extracted = extract_page_images(sgk_path, page, output_path.parent / "extracted_images")
        if extracted and target:
            target.add_run().add_picture(extracted[0]["path"], width=Cm(8))
            if description:
                target.add_run(f"\nHình minh họa: {description} (SGK, tr.{page}).")
        else:
            note = fallback_note(description or anchor or "Bổ sung hình minh họa theo hoạt động", page or 1)
            if target:
                target.insert_paragraph_before(note)
            else:
                doc.add_paragraph(note)


def generate_docx(meta: LessonMeta, template_path: Path | None, output_path: Path,
                  convert_equations: bool = True, sgk_path: Path | None = None) -> dict:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if meta.header_only:
        doc = _header_only(meta)
    elif template_path and template_path.exists():
        doc = Document(template_path)
        _set_header(doc, meta)
        _prune_optional_blocks(doc, meta)
    else:
        doc = _build_structured(meta)
    image_requests = (meta.lesson_content or {}).get("images", [])
    if image_requests:
        _attach_requested_images(doc, image_requests, sgk_path, output_path)
    converted = convert_document_equations(doc, Path(__file__).with_name("MML2OMML.XSL")) if convert_equations else 0
    doc.save(output_path)
    return {"path": str(output_path), "converted_equations": converted,
            "template_based": bool(template_path and template_path.exists()),
            "equation_note": "Công thức LaTeX nhận diện được đã chuyển thành OMML." if converted else "Không có biểu thức LaTeX cần chuyển; OMML có sẵn trong mẫu được giữ nguyên."}
