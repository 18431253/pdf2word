import copy
import re
import fitz  # pymupdf
from pdf2docx import Converter
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import json


_CTRL_RE = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]')
_GARBAGE_RE = re.compile(
    r'[�■-◿☀-⛿✀-➿-]{2,}'
)
_REPEAT_PUNCT_RE = re.compile(r'([^\w\s一-鿿])\1{3,}')


def _clean(text: str) -> str:
    text = _CTRL_RE.sub('', text)
    text = _GARBAGE_RE.sub('', text)
    text = _REPEAT_PUNCT_RE.sub(r'\1\1', text)
    return text.strip()


def is_scanned(pdf_path: str) -> bool:
    with fitz.open(pdf_path) as doc:
        for page in doc:
            if page.get_text().strip():
                return False
    return True


def _get_ocr_reader():
    import easyocr
    return easyocr.Reader(["ch_sim", "ch_tra", "en"], gpu=False)


def _ocr_to_docx(pdf_path: str, out_path: str, out_doc=None):
    reader = _get_ocr_reader()
    doc = out_doc if out_doc is not None else Document()
    with fitz.open(pdf_path) as fitz_doc:
        for page_num, page in enumerate(fitz_doc):
            if page_num > 0:
                doc.add_page_break()
            pix = page.get_pixmap(dpi=300)
            results = reader.readtext(pix.tobytes("png"), detail=0, paragraph=True)
            for line in results:
                line = _clean(line)
                if line:
                    doc.add_paragraph(line)
    if out_doc is None:
        doc.save(out_path)


def pdf_to_blocks(pdf_path: str) -> dict:
    """
    Parse PDF into structured blocks for preview.
    Returns { pages: [ { blocks: [ {type, text, size, bold, italic, align, color} ] } ] }
    """
    if is_scanned(pdf_path):
        return {"scanned": True, "pages": []}

    with fitz.open(pdf_path) as fitz_doc:
        # Compute average font size for heading detection
        all_sizes = []
        for page in fitz_doc:
            for block in page.get_text("dict")["blocks"]:
                if block["type"] != 0:
                    continue
                for line in block["lines"]:
                    for span in line["spans"]:
                        if span["text"].strip():
                            all_sizes.append(span["size"])

        avg_size = sum(all_sizes) / len(all_sizes) if all_sizes else 12
        h1_thresh = avg_size * 1.4
        h2_thresh = avg_size * 1.15

        pages = []
        for page in fitz_doc:
            page_blocks = []
            for block in page.get_text("dict")["blocks"]:
                if block["type"] != 0:
                    continue
                for line in block["lines"]:
                    spans = line["spans"]
                    if not spans:
                        continue

                    runs = []
                    for span in spans:
                        t = _clean(span["text"])
                        if not t:
                            continue
                        flags = span.get("flags", 0)
                        color = span.get("color", 0)
                        r = (color >> 16) & 0xff
                        g = (color >> 8) & 0xff
                        b = color & 0xff
                        runs.append({
                            "text": t,
                            "size": round(span["size"], 1),
                            "bold": bool("Bold" in span.get("font", "") or (flags & 16)),
                            "italic": bool("Italic" in span.get("font", "") or (flags & 2)),
                            "color": f"#{r:02x}{g:02x}{b:02x}" if color else "#000000",
                        })

                    if not runs:
                        continue

                    first = runs[0]
                    size = first["size"]
                    bold = first["bold"]

                    if bold and size >= h1_thresh:
                        btype = "h1"
                    elif bold and size >= h2_thresh:
                        btype = "h2"
                    else:
                        btype = "p"

                    bbox = block.get("bbox", [0, 0, 0, 0])
                    page_w = page.rect.width
                    left_margin = bbox[0]
                    right_margin = page_w - bbox[2]
                    if abs(left_margin - right_margin) < 20:
                        align = "center"
                    elif right_margin < left_margin * 0.5:
                        align = "right"
                    else:
                        align = "left"

                    page_blocks.append({
                        "type": btype,
                        "runs": runs,
                        "align": align,
                        "size": size,
                    })

            pages.append({"blocks": page_blocks})

    return {"scanned": False, "pages": pages}


def blocks_to_docx(blocks_data: dict, out_path: str, ref_path: str = None):
    """Convert edited blocks JSON back to docx."""
    if ref_path:
        ref_doc = Document(ref_path)
        doc = Document()
        _copy_styles_from_ref(ref_doc, doc)
        _copy_page_layout(ref_doc, doc)
        # Remove default empty paragraph safely
        for p in list(doc.paragraphs):
            p._element.getparent().remove(p._element)
    else:
        doc = Document()
        for p in list(doc.paragraphs):
            p._element.getparent().remove(p._element)

    type_to_style = {"h1": "Heading 1", "h2": "Heading 2", "p": "Normal"}
    pages = blocks_data.get("pages", [])

    for page_idx, page_data in enumerate(pages):
        for block in page_data.get("blocks", []):
            btype = block.get("type", "p")
            align = block.get("align", "left")
            runs = block.get("runs", [])

            style_name = type_to_style.get(btype, "Normal")
            if not _has_style(doc, style_name):
                style_name = "Normal"

            para = doc.add_paragraph(style=style_name)

            align_map = {
                "center": WD_ALIGN_PARAGRAPH.CENTER,
                "right": WD_ALIGN_PARAGRAPH.RIGHT,
                "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
                "left": WD_ALIGN_PARAGRAPH.LEFT,
            }
            para.alignment = align_map.get(align, WD_ALIGN_PARAGRAPH.LEFT)

            for run_data in runs:
                run = para.add_run(run_data.get("text", ""))
                run.bold = run_data.get("bold", False)
                run.italic = run_data.get("italic", False)
                size = run_data.get("size")
                if size:
                    run.font.size = Pt(size)
                color = run_data.get("color", "#000000")
                if color and color != "#000000":
                    try:
                        r = int(color[1:3], 16)
                        g = int(color[3:5], 16)
                        b = int(color[5:7], 16)
                        run.font.color.rgb = RGBColor(r, g, b)
                    except Exception:
                        pass

        # Page break between pages (except last)
        if page_idx < len(pages) - 1:
            doc.add_page_break()

    doc.save(out_path)


def convert_mode1(pdf_path: str, out_path: str):
    if is_scanned(pdf_path):
        _ocr_to_docx(pdf_path, out_path)
    else:
        cv = Converter(pdf_path)
        try:
            cv.convert(out_path, start=0, end=None)
        finally:
            cv.close()


def _extract_text_blocks(pdf_path: str) -> list:
    data = pdf_to_blocks(pdf_path)
    blocks = []
    for i, page in enumerate(data.get("pages", [])):
        if i > 0:
            blocks.append({"is_page_break": True})
        for b in page.get("blocks", []):
            text = " ".join(r["text"] for r in b.get("runs", []))
            blocks.append({
                "text": text,
                "is_heading": b["type"] in ("h1", "h2"),
                "level": 1 if b["type"] == "h1" else 2,
                "is_page_break": False,
            })
    return blocks


def _copy_styles_from_ref(ref_doc, out_doc):
    ref_styles_el = ref_doc.styles.element
    out_styles_el = out_doc.styles.element
    for style in ref_styles_el.findall(qn("w:style")):
        style_id = style.get(qn("w:styleId"))
        # Remove existing style with same ID before replacing
        for existing in out_styles_el.findall(qn("w:style")):
            if existing.get(qn("w:styleId")) == style_id:
                out_styles_el.remove(existing)
                break
        out_styles_el.append(copy.deepcopy(style))


def _copy_page_layout(ref_doc, out_doc):
    ref_sec = ref_doc.sections[0]
    out_sec = out_doc.sections[0]
    for attr in ("page_width", "page_height", "left_margin", "right_margin",
                 "top_margin", "bottom_margin", "header_distance", "footer_distance"):
        try:
            setattr(out_sec, attr, getattr(ref_sec, attr))
        except Exception:
            pass


def _has_style(doc, style_name: str) -> bool:
    try:
        doc.styles[style_name]
        return True
    except KeyError:
        return False


def convert_mode2(pdf_path: str, out_path: str, ref_path: str):
    ref_doc = Document(ref_path)
    out_doc = Document()
    _copy_styles_from_ref(ref_doc, out_doc)
    _copy_page_layout(ref_doc, out_doc)
    for p in list(out_doc.paragraphs):
        p._element.getparent().remove(p._element)

    if is_scanned(pdf_path):
        _ocr_to_docx(pdf_path, out_path, out_doc=out_doc)
        return

    blocks = _extract_text_blocks(pdf_path)
    for b in blocks:
        if b.get("is_page_break"):
            out_doc.add_page_break()
            continue
        text = b["text"]
        if b["is_heading"]:
            style_name = f"Heading {b['level']}"
            if not _has_style(out_doc, style_name):
                style_name = "Normal"
            out_doc.add_paragraph(text, style=style_name)
        else:
            out_doc.add_paragraph(text, style="Normal" if _has_style(out_doc, "Normal") else None)

    out_doc.save(out_path)
