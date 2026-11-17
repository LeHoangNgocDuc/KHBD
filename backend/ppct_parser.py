from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path


DATE = r"\d{2}/\d{2}/\d{4}"


def _text_from_pdf(path: Path) -> str:
    import fitz
    with fitz.open(path) as pdf:
        return "\n".join(page.get_text("text") for page in pdf)


def _text_from_docx(path: Path) -> str:
    from docx import Document
    doc = Document(path)
    chunks = [p.text for p in doc.paragraphs]
    for table in doc.tables:
        chunks.extend(" | ".join(c.text for c in row.cells) for row in table.rows)
    return "\n".join(chunks)


def _text_from_xlsx(path: Path) -> str:
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    return "\n".join(" | ".join(str(v or "") for v in row) for ws in wb for row in ws.iter_rows(values_only=True))


def extract_text(path: Path) -> str:
    ext = path.suffix.lower()
    if ext == ".pdf":
        return _text_from_pdf(path)
    if ext == ".docx":
        return _text_from_docx(path)
    if ext == ".xlsx":
        return _text_from_xlsx(path)
    raise ValueError("PPCT chỉ hỗ trợ PDF, DOCX hoặc XLSX.")


def _week_context(lines: list[str], index: int) -> tuple[int | None, str, str]:
    start = max(0, index - 5)
    context = " ".join(lines[start:index + 1])
    m = re.search(rf"Tuần\s*(\d+)\s*\(?\s*({DATE})\s*-\s*({DATE})", context, re.I)
    if not m:
        m = re.search(r"Tuần\s*(\d+)", context, re.I)
    if not m:
        return None, "", ""
    if len(m.groups()) >= 3:
        return int(m.group(1)), m.group(2), m.group(3)
    return int(m.group(1)), "", ""


def _date_ranges(text: str) -> dict[int, tuple[str, str]]:
    lines = [re.sub(r"\s+", " ", x).strip() for x in text.splitlines()]
    found: dict[int, tuple[str, str]] = {}
    for i, line in enumerate(lines):
        m = re.search(rf"Tuần\s*(\d+)\s*\(?\s*({DATE})\s*-\s*({DATE})", line, re.I)
        if m:
            found[int(m.group(1))] = (m.group(2), m.group(3))
        elif re.search(r"Tuần\s*\d+", line, re.I):
            m2 = re.search(rf"({DATE})\s*-\s*({DATE})", " ".join(lines[i:i + 4]))
            w = re.search(r"Tuần\s*(\d+)", line, re.I)
            if m2 and w:
                found[int(w.group(1))] = (m2.group(1), m2.group(2))
    return found


def parse_ppct(path: Path) -> dict:
    text = extract_text(path)
    lines = [re.sub(r"\s+", " ", x).strip() for x in text.splitlines()]
    ranges = _date_ranges(text)
    rows: list[dict] = []
    current_week: int | None = None
    for i, line in enumerate(lines):
        wm = re.search(r"Tuần\s*(\d+)", line, re.I)
        if wm:
            current_week = int(wm.group(1))
        # Table PDF extraction frequently puts the PPCT number and lesson on adjacent lines.
        pm = re.match(r"^(\d{1,3})\s+(.+)$", line)
        if pm:
            period = int(pm.group(1))
            title = pm.group(2).strip(" |-")
        elif re.fullmatch(r"\d{1,3}", line) and i + 1 < len(lines):
            period = int(line)
            title = lines[i + 1].strip(" |-")
        else:
            continue
        if len(title) < 5 and i + 1 < len(lines):
            title = (title + " " + lines[i + 1]).strip()
        if not re.search(r"[A-Za-zÀ-ỹ]{3}", title):
            continue
        # Week can be inherited from nearby PDF lines, otherwise PPCT documents often
        # have a dedicated week label before the numbered rows.
        week = current_week
        if week is None:
            week, _, _ = _week_context(lines, i)
        stream = "HH" if re.search(r"Tỉ số lượng giác|tam giác vuông|đường tròn|hình nón|hình trụ|hình cầu", title, re.I) else "DS"
        dates = ranges.get(week, ("", "")) if week else ("", "")
        rows.append({"week": week, "date_start": dates[0], "date_end": dates[1],
                     "period": period, "stream": stream, "title": title,
                     "source_line": i + 1})
    # De-duplicate repeated wrapped/table header detections while preserving order.
    unique = {}
    for row in rows:
        unique[(row["week"], row["period"], row["title"])] = row
    rows = list(unique.values())
    rows.sort(key=lambda r: (r["week"] or 999, r["period"]))
    return {"filename": path.name, "rows": rows,
            "warnings": [] if rows else ["Không nhận dạng được dòng PPCT; PDF có thể là bản scan hoặc bảng có bố cục khác."]}
