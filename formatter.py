"""
DocFormatter Engine — applies user-defined formatting rules to .docx documents.

Strategy:
1. Parse formatting rules from the UI (font, size, spacing, etc.)
2. For each uploaded .docx, preserve text content but enforce those rules.
3. AGGRESSIVE formatting:
   - Force consistent font across entire document
   - Force specified line spacing everywhere
   - Strip ALL colors (font color, highlight, shading)
   - Apply paragraph spacing rules
"""

from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_LINE_SPACING
from docx.oxml.ns import qn
import re
import io


BLACK = RGBColor(0, 0, 0)

# Default rules
DEFAULT_RULES = {
    "font_name": "Arial",
    "body_size": 12,
    "heading1_size": 16,
    "heading2_size": 14,
    "title_size": 20,
    "line_spacing": 1.5,
    "remove_colors": True,
    "remove_bold": False,
    "remove_italic": False,
}


def _get_heading_level(style_name):
    """Extract heading level from style name."""
    match = re.search(r"(\d)", style_name)
    if match:
        return int(match.group(1))
    if "heading" in style_name.lower() or "überschrift" in style_name.lower():
        return 1
    return None


def classify_paragraph(para):
    """Classify a paragraph as title, heading (with level), or body."""
    style_name = (para.style.name if para.style else "").lower()

    if "title" in style_name or "titel" in style_name:
        return ("title", None)

    if "heading" in style_name or "überschrift" in style_name:
        level = _get_heading_level(style_name)
        return ("heading", level or 1)

    # Heuristic: large bold text without heading style
    if para.runs and para.text.strip():
        run = para.runs[0]
        size = run.font.size
        if size and size >= Pt(16) and run.font.bold:
            return ("title", None)
        if size and size >= Pt(13) and run.font.bold:
            return ("heading", 1)

    return ("body", None)


def _strip_color_from_run(run):
    """Aggressively remove ALL color formatting from a run."""
    run.font.color.rgb = BLACK

    # Remove theme color
    rPr = run._element.find(qn("w:rPr"))
    if rPr is not None:
        color_elem = rPr.find(qn("w:color"))
        if color_elem is not None:
            for attr in ["themeColor", "themeTint", "themeShade"]:
                key = qn("w:" + attr)
                if key in color_elem.attrib:
                    del color_elem.attrib[key]

    # Remove highlight
    run.font.highlight_color = None

    # Remove run shading
    if rPr is not None:
        shd = rPr.find(qn("w:shd"))
        if shd is not None:
            rPr.remove(shd)


def _strip_para_shading(para):
    """Remove background/shading from paragraph level."""
    pPr = para._element.find(qn("w:pPr"))
    if pPr is not None:
        shd = pPr.find(qn("w:shd"))
        if shd is not None:
            pPr.remove(shd)


def _set_font_on_run(run, font_name):
    """Set font name on all font slots of a run."""
    run.font.name = font_name
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = run._element.makeelement(qn("w:rFonts"), {})
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), font_name)
    rFonts.set(qn("w:hAnsi"), font_name)
    rFonts.set(qn("w:eastAsia"), font_name)
    rFonts.set(qn("w:cs"), font_name)


def apply_rules_to_para(para, rules, para_type="body", heading_level=None):
    """Apply formatting rules to a paragraph."""
    pf = para.paragraph_format

    # Line spacing — always force
    pf.line_spacing = rules["line_spacing"]
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE

    # Strip paragraph shading
    if rules["remove_colors"]:
        _strip_para_shading(para)

    # Determine font size based on paragraph type
    if para_type == "title":
        target_size = Pt(rules["title_size"])
    elif para_type == "heading":
        if heading_level and heading_level >= 2:
            target_size = Pt(rules["heading2_size"])
        else:
            target_size = Pt(rules["heading1_size"])
    else:
        target_size = Pt(rules["body_size"])

    # Apply to all runs
    for run in para.runs:
        # Font
        _set_font_on_run(run, rules["font_name"])

        # Size
        run.font.size = target_size

        # Bold handling
        if rules["remove_bold"] and para_type == "body":
            run.font.bold = False
        elif para_type in ("title", "heading"):
            run.font.bold = True  # Headings always bold

        # Italic handling
        if rules["remove_italic"]:
            run.font.italic = False

        # Color stripping
        if rules["remove_colors"]:
            _strip_color_from_run(run)


def format_document(source_path, rules=None):
    """
    Apply formatting rules to a document.
    Returns bytes of the formatted .docx.
    """
    if rules is None:
        rules = DEFAULT_RULES.copy()

    doc = Document(source_path)

    # Apply paragraph formatting
    for para in doc.paragraphs:
        ptype, level = classify_paragraph(para)

        if not para.text.strip():
            # Empty paragraphs get body formatting for consistent spacing
            apply_rules_to_para(para, rules, "body")
            continue

        apply_rules_to_para(para, rules, ptype, level)

    # Also format text in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    apply_rules_to_para(para, rules, "body")

    # Save to bytes
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()
