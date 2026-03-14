"""
DocFormatter Engine — applies compiled formatting rules to .docx documents.

Rebuilds documents from scratch for maximum formatting control.
Rules are compiled from natural language by an LLM (see app.py).
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_LINE_SPACING, WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import re
import io
from datetime import datetime


def hex_to_rgb(hex_color):
    """Convert #RRGGBB to RGBColor."""
    hex_color = hex_color.lstrip("#")
    if len(hex_color) != 6:
        return RGBColor(0, 0, 0)
    try:
        return RGBColor(
            int(hex_color[0:2], 16),
            int(hex_color[2:4], 16),
            int(hex_color[4:6], 16),
        )
    except (ValueError, IndexError):
        return RGBColor(0, 0, 0)


ALIGNMENT_MAP = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


def _set_font_on_run(run, font_name):
    """Set font name on all font slots."""
    run.font.name = font_name
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = run._element.makeelement(qn("w:rFonts"), {})
        rPr.insert(0, rFonts)
    for attr in ["ascii", "hAnsi", "eastAsia", "cs"]:
        rFonts.set(qn(f"w:{attr}"), font_name)


def _apply_run_format(run, fmt, defaults):
    """Apply formatting to a single run."""
    font_name = fmt.get("font", defaults.get("font", "Arial"))
    _set_font_on_run(run, font_name)

    size = fmt.get("size", defaults.get("size", 12))
    run.font.size = Pt(size)

    if "bold" in fmt:
        run.font.bold = fmt["bold"]
    else:
        run.font.bold = defaults.get("bold", False)

    if "italic" in fmt:
        run.font.italic = fmt["italic"]
    else:
        run.font.italic = defaults.get("italic", False)

    color = fmt.get("color", defaults.get("color", "#000000"))
    run.font.color.rgb = hex_to_rgb(color)

    # Clean up highlights and shading
    run.font.highlight_color = None


def _apply_para_format(para, fmt, defaults):
    """Apply paragraph-level formatting."""
    pf = para.paragraph_format

    ls = fmt.get("line_spacing", defaults.get("line_spacing", 1.5))
    pf.line_spacing = ls
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE

    alignment = fmt.get("alignment", defaults.get("alignment", "left"))
    if alignment in ALIGNMENT_MAP:
        pf.alignment = ALIGNMENT_MAP[alignment]

    if "indent_cm" in fmt:
        pf.left_indent = Cm(fmt["indent_cm"])

    if "space_before_pt" in fmt:
        pf.space_before = Pt(fmt["space_before_pt"])
    if "space_after_pt" in fmt:
        pf.space_after = Pt(fmt["space_after_pt"])


def _add_page_numbers(doc, rules):
    """Add page numbers to footer."""
    pn = rules.get("page_numbers", {})
    if not pn.get("enabled", False):
        return

    section = doc.sections[0]
    footer = section.footer
    footer.is_linked_to_previous = False

    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()

    alignment = pn.get("alignment", "right")
    if alignment in ALIGNMENT_MAP:
        p.alignment = ALIGNMENT_MAP[alignment]

    # Add PAGE field
    run = p.add_run()
    fld_char_begin = OxmlElement("w:fldChar")
    fld_char_begin.set(qn("w:fldCharType"), "begin")
    run._element.append(fld_char_begin)

    run2 = p.add_run()
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    run2._element.append(instr)

    run3 = p.add_run()
    fld_char_end = OxmlElement("w:fldChar")
    fld_char_end.set(qn("w:fldCharType"), "end")
    run3._element.append(fld_char_end)

    # Format
    font_name = rules.get("defaults", {}).get("font", "Arial")
    size = pn.get("size", 9)
    for r in [run, run2, run3]:
        r.font.size = Pt(size)
        _set_font_on_run(r, font_name)
        r.font.color.rgb = RGBColor(0, 0, 0)


def _apply_inline_rules(text, inline_rules):
    """Split text into segments based on inline patterns.

    Returns list of {"text": str, "inline_fmt": dict} segments.
    """
    if not inline_rules:
        return [{"text": text, "inline_fmt": {}}]

    # Find all inline matches
    matches = []
    for rule in inline_rules:
        pattern = rule.get("pattern", "")
        if not pattern:
            continue
        try:
            for m in re.finditer(pattern, text):
                matches.append((m.start(), m.end(), rule.get("format", {})))
        except re.error:
            continue

    if not matches:
        return [{"text": text, "inline_fmt": {}}]

    # Sort by start position, resolve overlaps (first match wins)
    matches.sort(key=lambda x: x[0])
    filtered = []
    last_end = 0
    for start, end, fmt in matches:
        if start >= last_end:
            filtered.append((start, end, fmt))
            last_end = end

    # Build segments
    segments = []
    pos = 0
    for start, end, fmt in filtered:
        if start > pos:
            segments.append({"text": text[pos:start], "inline_fmt": {}})
        segments.append({"text": text[start:end], "inline_fmt": fmt})
        pos = end
    if pos < len(text):
        segments.append({"text": text[pos:], "inline_fmt": {}})

    return segments


def _match_paragraph_rule(text, rule, first_match_tracker):
    """Check if text matches a paragraph rule."""
    pattern = rule.get("pattern", "")
    match_type = rule.get("match_type", "regex")
    flags = re.IGNORECASE if rule.get("case_insensitive", True) else 0

    try:
        if match_type == "first_contains":
            rule_id = rule.get("name", pattern)
            if rule_id in first_match_tracker:
                return False
            if re.search(pattern, text, flags):
                first_match_tracker.add(rule_id)
                return True
        elif match_type == "contains":
            return bool(re.search(pattern, text, flags))
        elif match_type == "starts_with":
            return bool(re.match(pattern, text, flags))
        elif match_type == "regex":
            return bool(re.search(pattern, text, flags))
    except re.error:
        return False

    return False


def _find_character_config(character_name, dialogue_rules):
    """Find character-specific config, handling case-insensitive matching."""
    char_colors = dialogue_rules.get("character_colors", {})

    # Exact match (case-insensitive)
    for name, config in char_colors.items():
        if name.upper() == character_name.upper():
            return config

    # No match — assign from defaults
    return None


def _detect_scene_paragraphs(paragraphs_text, compiled_rules):
    """Find indices of scene heading paragraphs."""
    scene_indices = []
    for rule in compiled_rules.get("paragraph_rules", []):
        name = rule.get("name", "").lower()
        if "szene" in name or "scene" in name:
            for i, text in enumerate(paragraphs_text):
                try:
                    if re.search(rule["pattern"], text, re.IGNORECASE):
                        scene_indices.append(i)
                except re.error:
                    pass
            break
    return scene_indices


def format_document(source_path, compiled_rules):
    """
    Apply compiled rules to a document by rebuilding it from scratch.

    Args:
        source_path: Path to source .docx
        compiled_rules: Dict with compiled formatting rules

    Returns:
        Bytes of the formatted .docx
    """
    defaults = compiled_rules.get("defaults", {
        "font": "Arial", "size": 12, "color": "#000000",
        "bold": False, "italic": False, "alignment": "left", "line_spacing": 1.5
    })

    paragraph_rules = compiled_rules.get("paragraph_rules", [])
    dialogue_rules = compiled_rules.get("dialogue_rules", {})
    inline_rules = compiled_rules.get("inline_rules", [])
    scene_blank_lines = compiled_rules.get("scene_blank_lines", 0)

    # Read source document
    source = Document(source_path)
    source_texts = [p.text for p in source.paragraphs]

    # Detect scene boundaries for spacing
    scene_indices = _detect_scene_paragraphs(source_texts, compiled_rules) if scene_blank_lines > 0 else []

    # Create new document
    doc = Document()

    # Set default style
    style = doc.styles["Normal"]
    style.font.name = defaults.get("font", "Arial")
    style.font.size = Pt(defaults.get("size", 12))
    style.font.color.rgb = hex_to_rgb(defaults.get("color", "#000000"))
    style.paragraph_format.line_spacing = defaults.get("line_spacing", 1.5)
    style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE

    # Page numbers
    _add_page_numbers(doc, compiled_rules)

    # Track first_match rules
    first_match_tracker = set()

    # Track secondary character colors (for unknown characters)
    secondary_colors = dialogue_rules.get("secondary_shades", ["#0000FF", "#4169E1", "#1E90FF", "#00BFFF", "#6495ED"])
    unknown_chars = {}  # name -> assigned color
    secondary_idx = 0

    for i, text in enumerate(source_texts):
        # Add extra blank lines before scenes
        if i in scene_indices and i > 0:
            for _ in range(scene_blank_lines):
                blank_p = doc.add_paragraph()
                _apply_para_format(blank_p, {}, defaults)

        # Empty paragraph
        if not text.strip():
            p = doc.add_paragraph()
            _apply_para_format(p, {}, defaults)
            continue

        # --- Check paragraph rules ---
        matched_rule = None
        for rule in paragraph_rules:
            if _match_paragraph_rule(text, rule, first_match_tracker):
                matched_rule = rule
                break

        if matched_rule:
            fmt = matched_rule.get("format", {})
            para_fmt = {**fmt}

            p = doc.add_paragraph()
            _apply_para_format(p, para_fmt, defaults)

            # Process text with inline rules
            display_text = text.upper() if fmt.get("uppercase") else text
            segments = _apply_inline_rules(display_text, inline_rules)

            for seg in segments:
                run_fmt = {**defaults, **fmt, **seg["inline_fmt"]}
                if seg["inline_fmt"].get("uppercase"):
                    seg["text"] = seg["text"].upper()
                run = p.add_run(seg["text"])
                _apply_run_format(run, run_fmt, defaults)

            continue

        # --- Check dialogue rules ---
        if dialogue_rules.get("enabled", False):
            det_pattern = dialogue_rules.get(
                "detection_pattern",
                r"^([A-ZÄÖÜẞ][A-ZÄÖÜẞ\s.\-]+?)\s*[:：]"
            )
            try:
                m = re.match(det_pattern, text)
            except re.error:
                m = None

            if m:
                character_name = m.group(1).strip()
                name_text = text[: m.end()]  # "HEDDA:" or "DR. KHOURY:"
                dialogue_text = text[m.end() :]  # everything after colon

                # Find character config
                char_config = _find_character_config(character_name, dialogue_rules)

                if char_config is None:
                    # Unknown character — assign from secondary colors
                    upper_name = character_name.upper()
                    if upper_name not in unknown_chars:
                        # Check if there's a case_character_color
                        case_color = dialogue_rules.get("case_character_color")
                        if case_color and secondary_idx == 0:
                            unknown_chars[upper_name] = case_color
                        else:
                            color_idx = secondary_idx % len(secondary_colors) if secondary_colors else 0
                            unknown_chars[upper_name] = secondary_colors[color_idx] if secondary_colors else "#0000FF"
                        secondary_idx += 1
                    char_config = {"name_color": unknown_chars.get(upper_name, dialogue_rules.get("default_color", "#0000FF"))}

                name_format = dialogue_rules.get("name_format", {})
                name_color = char_config.get("name_color", defaults.get("color", "#000000"))

                # Paragraph-level formatting
                para_fmt = {}
                if char_config.get("text_indent_cm"):
                    para_fmt["indent_cm"] = char_config["text_indent_cm"]
                p = doc.add_paragraph()
                _apply_para_format(p, para_fmt, defaults)

                # Name run
                display_name = name_text.upper() if name_format.get("uppercase") else name_text
                name_run = p.add_run(display_name)
                _apply_run_format(name_run, {
                    **defaults,
                    "bold": name_format.get("bold", True),
                    "color": name_color,
                    "size": defaults.get("size", 12),
                }, defaults)

                # Dialogue text — process with inline rules
                text_italic = char_config.get("text_italic", defaults.get("italic", False))
                text_color = char_config.get("text_color", defaults.get("color", "#000000"))
                text_base_fmt = {
                    **defaults,
                    "italic": text_italic,
                    "color": text_color,
                    "bold": False,
                }

                segments = _apply_inline_rules(dialogue_text, inline_rules)
                for seg in segments:
                    seg_text = seg["text"]
                    seg_fmt = {**text_base_fmt, **seg["inline_fmt"]}
                    if seg["inline_fmt"].get("uppercase"):
                        seg_text = seg_text.upper()
                    run = p.add_run(seg_text)
                    _apply_run_format(run, seg_fmt, defaults)

                continue

        # --- Default paragraph (no rule matched) ---
        p = doc.add_paragraph()
        _apply_para_format(p, {}, defaults)

        segments = _apply_inline_rules(text, inline_rules)
        for seg in segments:
            seg_text = seg["text"]
            seg_fmt = {**defaults, **seg["inline_fmt"]}
            if seg["inline_fmt"].get("uppercase"):
                seg_text = seg_text.upper()
            run = p.add_run(seg_text)
            _apply_run_format(run, seg_fmt, defaults)

    # Save
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


def get_output_filename(original_name, compiled_rules):
    """Generate output filename with optional date suffix."""
    base = original_name.replace(".docx", "")

    if compiled_rules.get("filename_suffix_date"):
        fmt = compiled_rules.get("filename_date_format", "DDMM")
        now = datetime.now()
        if fmt == "DDMM":
            suffix = now.strftime("%d%m")
        elif fmt == "MMDD":
            suffix = now.strftime("%m%d")
        elif fmt == "YYYYMMDD":
            suffix = now.strftime("%Y%m%d")
        else:
            suffix = now.strftime("%d%m")
        return f"{base}_{suffix}.docx"

    return f"{base}_formatted.docx"
