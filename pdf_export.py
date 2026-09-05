import io
import os
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from pypdf import PdfReader, PdfWriter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas


WINDOWS_FONTS = Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"
BUNDLED_FONT = (
    Path(__file__).resolve().parent / "assets" / "fonts" / "NotoSansSC-Regular.ttf"
)

# Always use the bundled Unicode font. Hosted Linux environments do not include
# Windows fonts, and falling back to Times-Roman renders Chinese as black squares.
FONT_FILES = {
    "微软雅黑": BUNDLED_FONT,
    "黑体": BUNDLED_FONT,
    "华文楷体": BUNDLED_FONT,
    "楷体": BUNDLED_FONT,
    "楷体_GB2312": BUNDLED_FONT,
    "宋体": BUNDLED_FONT,
    "新宋体": BUNDLED_FONT,
    "华文宋体": BUNDLED_FONT,
}

FALLBACK_FONTS = [
    BUNDLED_FONT,
    WINDOWS_FONTS / "msyh.ttc",
    WINDOWS_FONTS / "msyhbd.ttc",
    WINDOWS_FONTS / "simhei.ttf",
    WINDOWS_FONTS / "simsun.ttc",
]

_registered = {}


def _register_font(path):
    path = Path(path)
    key = str(path).lower()
    if key in _registered:
        return _registered[key]
    if not path.exists():
        return None
    name = f"certfont_{len(_registered)}"
    if path.suffix.lower() == ".ttc":
        pdfmetrics.registerFont(TTFont(name, str(path), subfontIndex=0))
    else:
        pdfmetrics.registerFont(TTFont(name, str(path)))
    _registered[key] = name
    return name


def _default_font():
    for path in FALLBACK_FONTS:
        name = _register_font(path)
        if name:
            return name
    return "Times-Roman"


def _font_for_name(word_name):
    if word_name:
        path = FONT_FILES.get(word_name)
        if path:
            registered = _register_font(path)
            if registered:
                return registered
        for label, path in FONT_FILES.items():
            if label in word_name:
                registered = _register_font(path)
                if registered:
                    return registered
    return _default_font()


def _attr(value, name, default=None):
    if value is None:
        return default
    return getattr(value, name, default)


def _run_font_name(run):
    rPr = run._element.find(qn("w:rPr"))
    if rPr is not None:
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is not None:
            for attr in ("w:eastAsia", "w:ascii", "w:hAnsi", "w:cs"):
                name = rFonts.get(qn(attr))
                if name:
                    return name
    return run.font.name


def _run_size_pt(run, paragraph):
    if run.font.size:
        return run.font.size.pt
    pPr = paragraph._p.find(qn("w:pPr"))
    if pPr is not None:
        rPr = pPr.find(qn("w:rPr"))
        if rPr is not None:
            size = rPr.find(qn("w:sz"))
            if size is not None and size.get(qn("w:val")):
                return int(size.get(qn("w:val"))) / 2
    return 14


def _run_bold(run):
    return bool(run.font.bold)


def _run_underline(run):
    return bool(run.font.underline)


def _paragraph_line_height(paragraph, fallback_size):
    fmt = paragraph.paragraph_format
    spacing = fmt.line_spacing
    rule = fmt.line_spacing_rule
    if spacing is None:
        return fallback_size * 1.2
    if rule == WD_LINE_SPACING.EXACTLY:
        return max(_attr(spacing, "pt", fallback_size), fallback_size)
    if isinstance(spacing, (int, float)):
        return fallback_size * float(spacing)
    return _attr(spacing, "pt", fallback_size * 1.2)


def _iter_run_chunks(run, paragraph):
    text = run.text or ""
    if not text:
        return
    yield {
        "text": text,
        "font": _font_for_name(_run_font_name(run)),
        "size": _run_size_pt(run, paragraph),
        "bold": _run_bold(run),
        "underline": _run_underline(run),
    }


def docx_bytes_to_pdf(docx_bytes):
    """Render a single-page (or few-page) DOCX certificate to PDF bytes."""
    document = Document(io.BytesIO(docx_bytes))
    section = document.sections[0]
    page_width = section.page_width.pt
    page_height = section.page_height.pt
    left = section.left_margin.pt
    right = section.right_margin.pt
    top = section.top_margin.pt
    bottom = section.bottom_margin.pt
    content_width = page_width - left - right

    buffer = io.BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=(page_width, page_height))
    y = page_height - top
    default_font = _default_font()

    for paragraph in document.paragraphs:
        chunks = []
        for run in paragraph.runs:
            chunks.extend(_iter_run_chunks(run, paragraph))
        max_size = max((chunk["size"] for chunk in chunks), default=14)
        line_height = _paragraph_line_height(paragraph, max_size)
        space_before = _attr(paragraph.paragraph_format.space_before, "pt", 0) or 0
        first_indent = _attr(paragraph.paragraph_format.first_line_indent, "pt", 0) or 0
        align = paragraph.alignment
        y -= space_before

        if not chunks:
            y -= line_height
            if y < bottom:
                pdf.showPage()
                y = page_height - top
            continue

        lines = _wrap_chunks(chunks, content_width, first_indent, default_font)
        for line_index, line in enumerate(lines):
            if y - line_height < bottom - 1:
                pdf.showPage()
                y = page_height - top
            indent = first_indent if line_index == 0 else 0
            _draw_line(pdf, line, left, y - line_height * 0.78, content_width, indent, align)
            y -= line_height

    pdf.save()
    return buffer.getvalue()


def _wrap_chunks(chunks, content_width, first_indent, default_font):
    lines = []
    current = []
    width_used = first_indent

    def flush():
        nonlocal current, width_used
        if current:
            lines.append(current)
        current = []
        width_used = 0

    for chunk in chunks:
        font = chunk["font"] or default_font
        size = chunk["size"]
        for char in chunk["text"]:
            if char == "\t":
                char = "    "
            if char == "\n":
                flush()
                continue
            char_width = pdfmetrics.stringWidth(char, font, size)
            available = content_width - width_used
            if current and char_width > available:
                flush()
            current.append(
                {
                    "text": char,
                    "font": font,
                    "size": size,
                    "bold": chunk["bold"],
                    "underline": chunk["underline"],
                    "width": char_width,
                }
            )
            width_used += char_width
    flush()
    return lines or [[]]


def _draw_line(pdf, line, left, baseline, content_width, indent, align):
    total = indent + sum(item["width"] for item in line)
    if align == WD_ALIGN_PARAGRAPH.CENTER:
        x = left + max((content_width - total) / 2, 0)
    elif align == WD_ALIGN_PARAGRAPH.RIGHT:
        x = left + max(content_width - total, 0)
    else:
        x = left + indent

    for item in _coalesce_line(line):
        pdf.setFont(item["font"], item["size"])
        pdf.setFillGray(0)
        pdf.drawString(x, baseline, item["text"])
        if item["underline"]:
            pdf.setStrokeGray(0)
            pdf.setLineWidth(0.7)
            pdf.line(x, baseline - 1.2, x + item["width"], baseline - 1.2)
        x += item["width"]


def _coalesce_line(line):
    merged = []
    for item in line:
        if (
            merged
            and merged[-1]["font"] == item["font"]
            and merged[-1]["size"] == item["size"]
            and merged[-1]["underline"] == item["underline"]
        ):
            merged[-1]["text"] += item["text"]
            merged[-1]["width"] += item["width"]
        else:
            merged.append(dict(item))
    return merged


def merge_pdfs(pdf_pages):
    """Concatenate one-page PDFs into a single document."""
    writer = PdfWriter()
    for data in pdf_pages:
        reader = PdfReader(io.BytesIO(data))
        for page in reader.pages:
            writer.add_page(page)
    output = io.BytesIO()
    writer.write(output)
    return output.getvalue()
