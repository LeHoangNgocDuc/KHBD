from __future__ import annotations

import re
from pathlib import Path
from lxml import etree

M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NS = {"m": M}
TAG_RE = re.compile(r"\$\$(.+?)\$\$|\$(.+?)\$|\\\((.+?)\\\)|\\\[(.+?)\\\]", re.S)


def _m(tag: str):
    return etree.Element(f"{{{M}}}{tag}")


def _run(text: str):
    r = _m("r")
    t = _m("t")
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    r.append(t)
    return r


def _append_flat(parent, node):
    # MathML mrow/semantics are grouping constructs; Word OMML uses their
    # contents directly except where a structure explicitly needs an m:e slot.
    if node is None:
        return
    if etree.QName(node).namespace == M and etree.QName(node).localname == "e":
        for child in list(node):
            parent.append(child)
    else:
        parent.append(node)


def _mathml_node(el):
    name = etree.QName(el).localname
    children = [c for c in el if isinstance(c.tag, str) and etree.QName(c).localname not in ("annotation",)]
    if name in {"math", "semantics", "mrow", "mstyle", "mpadded", "mphantom"}:
        parent = _m("oMath") if name == "math" else _m("e")
        for child in children:
            node = _mathml_node(child)
            _append_flat(parent, node)
        return parent
    if name in {"mi", "mn", "mo", "mtext", "ms"}:
        val = "".join(el.itertext())
        return _run(val)
    if name == "mfrac" and len(children) >= 2:
        node = _m("f"); node.append(_m("num")); node.append(_m("den"))
        for c in children[0]:
            x = _mathml_node(c)
            if x is not None: node[0].append(x)
        for c in children[1]:
            x = _mathml_node(c)
            if x is not None: node[1].append(x)
        return node
    if name in {"msup", "msub", "msubsup"} and len(children) >= 2:
        kind = {"msup": "sSup", "msub": "sSub", "msubsup": "sSubSup"}[name]
        node = _m(kind)
        for slot, child in zip(("e", "sup", "sub", "sup"), children):
            part = _m(slot)
            for c in child:
                x = _mathml_node(c)
                if x is not None: part.append(x)
            node.append(part)
        return node
    if name in {"msqrt", "mroot"} and children:
        node = _m("rad"); pr = _m("radPr"); deg = _m("deg"); base = _m("e")
        hide = _m("degHide"); hide.set("{http://schemas.openxmlformats.org/officeDocument/2006/math}val", "1" if name == "msqrt" else "0")
        pr.append(hide); node.append(pr)
        if name == "mroot" and len(children) > 1:
            for c in children[1]:
                x = _mathml_node(c)
                if x is not None: deg.append(x)
        node.append(deg)
        for c in children[0]:
            x = _mathml_node(c)
            if x is not None: base.append(x)
        node.append(base)
        return node
    if name in {"mover", "munder", "munderover"} and children:
        node = _m({"mover": "limUpp", "munder": "limLow", "munderover": "limUpp"}[name])
        base = _m("e")
        for c in children[0]:
            x = _mathml_node(c)
            if x is not None: base.append(x)
        node.append(base)
        if len(children) > 1:
            sup = _m("sup" if name != "munder" else "lim")
            for c in children[1]:
                x = _mathml_node(c)
                if x is not None: sup.append(x)
            node.append(sup)
        if name == "munderover" and len(children) > 2:
            upper = _m("sup")
            for c in children[2]:
                x = _mathml_node(c)
                if x is not None: upper.append(x)
            node.append(upper)
        return node
    if name in {"mtable", "mtr", "mtd"}:
        node = _m({"mtable": "m", "mtr": "mr", "mtd": "e"}[name])
        for child in children:
            x = _mathml_node(child)
            if x is not None: node.append(x)
        return node
    if name == "mfenced":
        text = "".join(el.itertext())
        return _run(text)
    if name in {"mspace", "annotation"}:
        return None
    if children:
        node = _m("e")
        for child in children:
            x = _mathml_node(child)
            if x is not None: node.append(x)
        return node
    return _run("".join(el.itertext()))


def _mathml_to_omml(mathml: str, xslt_path: Path | None = None):
    root = etree.fromstring(mathml.encode("utf-8"))
    if xslt_path and xslt_path.exists():
        transform = etree.XSLT(etree.parse(str(xslt_path)))
        result = transform(root)
        candidates = result.xpath("//*[local-name()='oMath']")
        if candidates:
            return candidates[0]
    node = _mathml_node(root)
    if node is None:
        raise ValueError("Không chuyển được biểu thức MathML thành OMML.")
    if etree.QName(node).localname != "oMath":
        wrapper = _m("oMath"); wrapper.append(node); node = wrapper
    return node


def convert_latex(latex: str, xslt_path: Path | None = None):
    try:
        from latex2mathml.converter import convert
    except ImportError as exc:
        raise RuntimeError("Thiếu latex2mathml. Cài các gói trong requirements.txt để biên dịch công thức.") from exc
    return _mathml_to_omml(convert(latex), xslt_path)


def convert_paragraph(paragraph, xslt_path: Path | None = None) -> int:
    text = paragraph.text
    matches = list(TAG_RE.finditer(text))
    if not matches:
        return 0
    # Keep paragraph style and surrounding text; replace only the run payload.
    p = paragraph._p
    for child in list(p):
        if etree.QName(child).localname != "pPr":
            p.remove(child)
    cursor = 0
    for match in matches:
        if match.start() > cursor:
            from docx.oxml import OxmlElement
            run = OxmlElement("w:r"); t = OxmlElement("w:t")
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            t.text = text[cursor:match.start()]; run.append(t); p.append(run)
        latex = next(g for g in match.groups() if g is not None)
        p.append(convert_latex(latex, xslt_path))
        cursor = match.end()
    if cursor < len(text):
        from docx.oxml import OxmlElement
        run = OxmlElement("w:r"); t = OxmlElement("w:t")
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = text[cursor:]; run.append(t); p.append(run)
    return len(matches)


def convert_document_equations(doc, xslt_path: Path | None = None) -> int:
    count = 0
    paragraphs = list(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                paragraphs.extend(cell.paragraphs)
    for paragraph in paragraphs:
        count += convert_paragraph(paragraph, xslt_path)
    return count
