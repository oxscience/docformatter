"""
DocFormatter Engine — applies compiled formatting rules to .docx documents.

Two-pass approach:
1. Classify every paragraph (title, episode, scene, character name, dialogue,
   stage direction, SFX/ATM, time marker, etc.)
2. Format based on classification + speaker context

Handles Hörspiel/screenplay format where character names are on SEPARATE lines
from their dialogue (not "NAME: text" on one line).
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


# ─── Text transformations ────────────────────────────────────────────────────

def _convert_brackets(text, compiled_rules):
    """Convert square brackets to round brackets if configured."""
    if compiled_rules.get("convert_brackets_to_round", False):
        text = text.replace("[", "(").replace("]", ")")
    return text


# ─── Low-level formatting helpers ────────────────────────────────────────────

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


def _format_run(run, font, size_pt, bold, italic, color_hex, underline=False):
    """Apply all formatting to a run."""
    _set_font_on_run(run, font)
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    run.font.underline = underline
    run.font.color.rgb = hex_to_rgb(color_hex)
    run.font.highlight_color = None


def _format_paragraph(para, alignment="left", line_spacing=1.5, indent_cm=None,
                      space_before_pt=None, space_after_pt=None):
    """Apply paragraph-level formatting."""
    pf = para.paragraph_format
    pf.line_spacing = line_spacing
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE

    if alignment in ALIGNMENT_MAP:
        pf.alignment = ALIGNMENT_MAP[alignment]

    if indent_cm is not None:
        pf.left_indent = Cm(indent_cm)

    if space_before_pt is not None:
        pf.space_before = Pt(space_before_pt)
    if space_after_pt is not None:
        pf.space_after = Pt(space_after_pt)


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

    run = p.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    run._element.append(fld_begin)

    run2 = p.add_run()
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    run2._element.append(instr)

    run3 = p.add_run()
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run3._element.append(fld_end)

    font = rules.get("defaults", {}).get("font", "Arial")
    size = pn.get("size", 9)
    for r in [run, run2, run3]:
        r.font.size = Pt(size)
        _set_font_on_run(r, font)
        r.font.color.rgb = RGBColor(0, 0, 0)


# ─── Classification pass ─────────────────────────────────────────────────────

# Paragraph types
T_EMPTY = "empty"
T_TITLE = "title"           # Series name (first occurrence)
T_EPISODE = "episode"       # "Folge X: ..."
T_SCENE = "scene"           # "SZENE 1: ..."
T_CHARACTER = "character"   # Standalone character name line
T_DIALOGUE = "dialogue"     # Dialogue text (follows character name)
T_STAGE_DIR = "stage_dir"   # Stage direction in () — own line
T_SFX_ATM = "sfx_atm"       # [SFX/ATM: ...], [TITELSONG: ...]
T_TIME = "time_marker"      # [Kumulierte Zeit: ...]
T_LEIT = "leit_objekt"      # [LEIT-OBJEKT ...]
T_BRACKET = "bracket_dir"   # Other [stage directions]
T_BODY = "body"             # Default


def _build_character_set(compiled_rules):
    """Extract all known character names from compiled rules."""
    names = set()
    dialogue = compiled_rules.get("dialogue_rules", {})
    for name in dialogue.get("character_colors", {}).keys():
        names.add(name.upper().strip())
    return names


def _is_character_name(text, known_characters):
    """Check if a paragraph is a standalone character name line."""
    stripped = text.strip()
    if not stripped:
        return False

    # Must be short (max ~4 words)
    words = stripped.split()
    if len(words) > 5:
        return False

    # Check against known characters (case-insensitive)
    if stripped.upper() in known_characters:
        return True

    # Heuristic: all uppercase, no punctuation except periods/spaces
    clean = stripped.replace(".", "").replace(" ", "").replace("-", "")
    if clean and clean.upper() == clean and clean.isalpha():
        upper = stripped.upper()
        if any(upper.startswith(x) for x in ["SZENE", "SFX", "ATM", "LEIT"]):
            return False
        return True

    return False


def classify_paragraphs(texts, compiled_rules):
    """
    First pass: classify every paragraph.

    Returns list of (type, speaker) tuples where speaker is the current
    character name for dialogue paragraphs.
    """
    known_characters = _build_character_set(compiled_rules)
    paragraph_rules = compiled_rules.get("paragraph_rules", [])

    classifications = []
    first_match_tracker = set()

    for i, text in enumerate(texts):
        stripped = text.strip()

        # Empty
        if not stripped:
            classifications.append((T_EMPTY, None))
            continue

        # --- Check paragraph_rules from compiled rules ---
        matched_rule = None
        for rule in paragraph_rules:
            pattern = rule.get("pattern", "")
            match_type = rule.get("match_type", "regex")
            flags = re.IGNORECASE if rule.get("case_insensitive", True) else 0

            try:
                if match_type == "first_contains":
                    rule_id = rule.get("name", pattern)
                    if rule_id not in first_match_tracker and re.search(pattern, stripped, flags):
                        first_match_tracker.add(rule_id)
                        matched_rule = rule
                        break
                elif match_type in ("contains", "starts_with", "regex"):
                    if re.search(pattern, stripped, flags):
                        matched_rule = rule
                        break
            except re.error:
                continue

        if matched_rule:
            name = matched_rule.get("name", "").lower()
            if "titel" in name or "serien" in name or "title" in name:
                classifications.append((T_TITLE, None))
            elif "folge" in name or "episode" in name:
                classifications.append((T_EPISODE, None))
            elif "szene" in name or "scene" in name:
                classifications.append((T_SCENE, None))
            elif "zeit" in name or "time" in name or "kumuliert" in name:
                classifications.append((T_TIME, None))
            else:
                classifications.append((f"rule:{matched_rule.get('name', '')}", None))
            continue

        # --- Hardcoded structural patterns ---

        # Time marker
        if re.match(r"^\[Kumulierte\s+Zeit", stripped, re.IGNORECASE) or \
           re.match(r"^\(Kumulierte\s+Zeit", stripped, re.IGNORECASE):
            classifications.append((T_TIME, None))
            continue

        # Scene heading
        if re.match(r"^SZENE\s+\d+|^Szene\s+\d+", stripped):
            classifications.append((T_SCENE, None))
            continue

        # Episode
        if re.match(r".*Folge\s+\d+", stripped, re.IGNORECASE):
            classifications.append((T_EPISODE, None))
            continue

        # SFX/ATM in brackets (both [] and ())
        if re.match(r"^[\[\(](SFX|ATM|TITELSONG|OUTRO)", stripped, re.IGNORECASE):
            classifications.append((T_SFX_ATM, None))
            continue

        # LEIT-OBJEKT in brackets
        if re.match(r"^[\[\(]LEIT-OBJEKT", stripped, re.IGNORECASE):
            classifications.append((T_LEIT, None))
            continue

        # Other bracket/paren directions (full line enclosed)
        if (stripped.startswith("[") and stripped.endswith("]")) or \
           (stripped.startswith("(") and stripped.endswith(")")):
            # Check if it's a stage direction (short, in parens)
            # or a bracket direction (longer, descriptive)
            if stripped.startswith("(") and stripped.endswith(")"):
                classifications.append((T_STAGE_DIR, None))
            else:
                classifications.append((T_BRACKET, None))
            continue

        # Character name (standalone line)
        if _is_character_name(stripped, known_characters):
            classifications.append((T_CHARACTER, stripped.upper()))
            continue

        # Default: body text
        classifications.append((T_BODY, None))

    # --- Second pass: assign speaker to dialogue ---
    result = []
    current_speaker = None

    for i, (ptype, data) in enumerate(classifications):
        if ptype == T_CHARACTER:
            current_speaker = data
            result.append((T_CHARACTER, data))
        elif ptype == T_EMPTY:
            result.append((T_EMPTY, current_speaker))
        elif ptype in (T_SCENE, T_TITLE, T_EPISODE):
            current_speaker = None
            result.append((ptype, None))
        elif ptype == T_BODY:
            if current_speaker:
                result.append((T_DIALOGUE, current_speaker))
            else:
                result.append((T_BODY, None))
        elif ptype == T_STAGE_DIR:
            result.append((T_STAGE_DIR, current_speaker))
        elif ptype in (T_SFX_ATM, T_BRACKET, T_LEIT):
            result.append((ptype, current_speaker))
        else:
            result.append((ptype, data))

    return result


# ─── Formatting pass ─────────────────────────────────────────────────────────

def _get_character_color(character_name, dialogue_rules, unknown_tracker):
    """Get the color config for a character."""
    char_colors = dialogue_rules.get("character_colors", {})

    for name, config in char_colors.items():
        if name.upper() == character_name.upper():
            return config

    upper = character_name.upper()
    if upper not in unknown_tracker:
        shades = dialogue_rules.get("secondary_shades", ["#0000FF", "#4169E1", "#1E90FF", "#00BFFF"])
        case_color = dialogue_rules.get("case_character_color")

        if case_color and len(unknown_tracker) == 0:
            unknown_tracker[upper] = {"name_color": case_color}
        elif shades:
            idx = len(unknown_tracker) % len(shades)
            unknown_tracker[upper] = {"name_color": shades[idx]}
        else:
            default = dialogue_rules.get("default_color", "#0000FF")
            unknown_tracker[upper] = {"name_color": default}

    return unknown_tracker.get(upper, {"name_color": "#0000FF"})


def _add_text_with_inline_formatting(para, text, base_font, base_size, base_bold,
                                      base_italic, base_color, inline_rules,
                                      base_underline=False):
    """Add text to paragraph, applying inline rules (e.g., italic for parentheses)."""
    if not inline_rules:
        run = para.add_run(text)
        _format_run(run, base_font, base_size, base_bold, base_italic, base_color, base_underline)
        return

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
        run = para.add_run(text)
        _format_run(run, base_font, base_size, base_bold, base_italic, base_color, base_underline)
        return

    # Sort and filter overlaps
    matches.sort(key=lambda x: x[0])
    filtered = []
    last_end = 0
    for start, end, fmt in matches:
        if start >= last_end:
            filtered.append((start, end, fmt))
            last_end = end

    # Build runs
    pos = 0
    for start, end, fmt in filtered:
        if start > pos:
            run = para.add_run(text[pos:start])
            _format_run(run, base_font, base_size, base_bold, base_italic, base_color, base_underline)

        seg_text = text[start:end]
        if fmt.get("uppercase"):
            seg_text = seg_text.upper()

        run = para.add_run(seg_text)
        _format_run(
            run, base_font,
            fmt.get("size", base_size),
            fmt.get("bold", base_bold),
            fmt.get("italic", base_italic),
            fmt.get("color", base_color),
            fmt.get("underline", base_underline),
        )
        pos = end

    if pos < len(text):
        run = para.add_run(text[pos:])
        _format_run(run, base_font, base_size, base_bold, base_italic, base_color, base_underline)


def format_document(source_path, compiled_rules):
    """
    Apply compiled rules to a document by rebuilding from scratch.
    Returns bytes of the formatted .docx.
    """
    defaults = compiled_rules.get("defaults", {
        "font": "Arial", "size": 12, "color": "#000000",
        "bold": False, "italic": False, "alignment": "left", "line_spacing": 1.5
    })

    d_font = defaults.get("font", "Arial")
    d_size = defaults.get("size", 12)
    d_color = defaults.get("color", "#000000")
    d_bold = defaults.get("bold", False)
    d_italic = defaults.get("italic", False)
    d_align = defaults.get("alignment", "left")
    d_spacing = defaults.get("line_spacing", 1.5)

    dialogue_rules = compiled_rules.get("dialogue_rules", {})
    inline_rules = compiled_rules.get("inline_rules", [])
    paragraph_rules = compiled_rules.get("paragraph_rules", [])
    scene_blank = compiled_rules.get("scene_blank_lines", 0)
    name_format = dialogue_rules.get("name_format", {"bold": True, "uppercase": True})

    # Spacing after character name (in pt). 0.5 line at 12pt ≈ 6pt
    name_space_after = compiled_rules.get("character_name_space_after_pt", 6)

    # Stage direction / SFX / bracket size (default to d_size, can be overridden)
    stage_dir_size = compiled_rules.get("stage_direction_size", d_size)

    # Read source
    source = Document(source_path)
    texts = [p.text for p in source.paragraphs]

    # Classify
    classifications = classify_paragraphs(texts, compiled_rules)

    # Create new document
    doc = Document()

    # Set default style
    style = doc.styles["Normal"]
    style.font.name = d_font
    style.font.size = Pt(d_size)
    style.font.color.rgb = hex_to_rgb(d_color)
    style.paragraph_format.line_spacing = d_spacing
    style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE

    # Page numbers
    _add_page_numbers(doc, compiled_rules)

    # Track unknown character colors
    unknown_chars = {}

    # Find rule formats by name for quick lookup
    rule_formats = {}
    for rule in paragraph_rules:
        rule_formats[rule.get("name", "")] = rule.get("format", {})

    for i, (ptype, speaker) in enumerate(classifications):
        text = texts[i]
        stripped = text.strip()

        # Apply bracket conversion
        stripped = _convert_brackets(stripped, compiled_rules)

        # --- EMPTY ---
        if ptype == T_EMPTY:
            p = doc.add_paragraph()
            _format_paragraph(p, d_align, d_spacing)
            continue

        # --- TITLE (series name) ---
        if ptype == T_TITLE:
            fmt = _find_rule_format("titel", paragraph_rules) or \
                  _find_rule_format("title", paragraph_rules) or \
                  _find_rule_format("serien", paragraph_rules)
            size = fmt.get("size", 20) if fmt else 20
            bold = fmt.get("bold", True) if fmt else True
            upper = fmt.get("uppercase", True) if fmt else True
            underline = fmt.get("underline", False) if fmt else False

            p = doc.add_paragraph()
            _format_paragraph(p, fmt.get("alignment", d_align) if fmt else d_align, d_spacing)
            display = stripped.upper() if upper else stripped
            run = p.add_run(display)
            _format_run(run, d_font, size, bold, False, d_color, underline)
            continue

        # --- EPISODE ---
        if ptype == T_EPISODE:
            fmt = _find_rule_format("folge", paragraph_rules) or \
                  _find_rule_format("episode", paragraph_rules)
            size = fmt.get("size", 16) if fmt else 16
            bold = fmt.get("bold", True) if fmt else True
            upper = fmt.get("uppercase", False) if fmt else False
            underline = fmt.get("underline", False) if fmt else False

            # Space before episode line (e.g., 0.5 line = 6pt)
            space_before = fmt.get("space_before_pt") if fmt else None

            p = doc.add_paragraph()
            _format_paragraph(p, d_align, d_spacing, space_before_pt=space_before)
            display = stripped.upper() if upper else stripped
            run = p.add_run(display)
            _format_run(run, d_font, size, bold, False, d_color, underline)
            continue

        # --- SCENE HEADING ---
        if ptype == T_SCENE:
            # Add blank lines before scene
            if scene_blank and i > 0:
                for _ in range(scene_blank):
                    bp = doc.add_paragraph()
                    _format_paragraph(bp, d_align, d_spacing)

            fmt = _find_rule_format("szene", paragraph_rules) or \
                  _find_rule_format("scene", paragraph_rules)
            size = fmt.get("size", 13) if fmt else 13
            bold = fmt.get("bold", True) if fmt else True
            upper = fmt.get("uppercase", True) if fmt else True
            underline = fmt.get("underline", False) if fmt else False

            p = doc.add_paragraph()
            _format_paragraph(p, d_align, d_spacing)
            display = stripped.upper() if upper else stripped
            run = p.add_run(display)
            _format_run(run, d_font, size, bold, False, d_color, underline)
            continue

        # --- TIME MARKER ---
        if ptype == T_TIME:
            fmt = _find_rule_format("zeit", paragraph_rules) or \
                  _find_rule_format("time", paragraph_rules) or \
                  _find_rule_format("kumuliert", paragraph_rules)
            size = fmt.get("size", 9) if fmt else 9
            italic = fmt.get("italic", True) if fmt else True
            align = fmt.get("alignment", "right") if fmt else "right"

            p = doc.add_paragraph()
            _format_paragraph(p, align, d_spacing)
            run = p.add_run(stripped)
            _format_run(run, d_font, size, False, italic, d_color)
            continue

        # --- CHARACTER NAME ---
        if ptype == T_CHARACTER:
            char_config = _get_character_color(speaker, dialogue_rules, unknown_chars)
            name_color = char_config.get("name_color", d_color)
            name_bold = name_format.get("bold", True)
            name_upper = name_format.get("uppercase", True)

            # Character name is NEVER indented — only dialogue text gets indent
            p = doc.add_paragraph()
            _format_paragraph(p, d_align, d_spacing, space_after_pt=name_space_after)
            display = stripped.upper() if name_upper else stripped
            run = p.add_run(display)
            _format_run(run, d_font, d_size, name_bold, False, name_color)
            continue

        # --- DIALOGUE ---
        if ptype == T_DIALOGUE and speaker:
            char_config = _get_character_color(speaker, dialogue_rules, unknown_chars)
            text_color = char_config.get("text_color", d_color)
            text_italic = char_config.get("text_italic", False)
            indent = char_config.get("text_indent_cm", None)

            p = doc.add_paragraph()
            _format_paragraph(p, d_align, d_spacing, indent_cm=indent)
            _add_text_with_inline_formatting(
                p, stripped, d_font, d_size, d_bold, text_italic, text_color, inline_rules
            )
            continue

        # --- STAGE DIRECTION (in parentheses, own line) ---
        if ptype == T_STAGE_DIR:
            # Stage directions get indent if current speaker has it
            indent = None
            if speaker:
                char_config = _get_character_color(speaker, dialogue_rules, unknown_chars)
                indent = char_config.get("text_indent_cm", None)

            p = doc.add_paragraph()
            _format_paragraph(p, d_align, d_spacing, indent_cm=indent)
            run = p.add_run(stripped)
            _format_run(run, d_font, stage_dir_size, False, True, d_color)
            continue

        # --- SFX/ATM ---
        if ptype == T_SFX_ATM:
            p = doc.add_paragraph()
            _format_paragraph(p, d_align, d_spacing)
            run = p.add_run(stripped)
            _format_run(run, d_font, stage_dir_size, False, True, d_color)
            continue

        # --- LEIT-OBJEKT ---
        if ptype == T_LEIT:
            p = doc.add_paragraph()
            _format_paragraph(p, d_align, d_spacing)

            leit_match = re.search(r"LEIT-OBJEKT|Leit-[Oo]bjekt", stripped, re.IGNORECASE)
            if leit_match:
                before = stripped[:leit_match.start()]
                leit_text = stripped[leit_match.start():leit_match.end()]
                after = stripped[leit_match.end():]

                if before:
                    run = p.add_run(before)
                    _format_run(run, d_font, stage_dir_size, False, True, d_color)

                run = p.add_run(leit_text.upper())
                _format_run(run, d_font, stage_dir_size, True, True, d_color)

                if after:
                    run = p.add_run(after)
                    _format_run(run, d_font, stage_dir_size, False, True, d_color)
            else:
                run = p.add_run(stripped)
                _format_run(run, d_font, stage_dir_size, False, True, d_color)
            continue

        # --- BRACKET DIRECTION ---
        if ptype == T_BRACKET:
            p = doc.add_paragraph()
            _format_paragraph(p, d_align, d_spacing)
            run = p.add_run(stripped)
            _format_run(run, d_font, stage_dir_size, False, True, d_color)
            continue

        # --- Matched paragraph rule ---
        if ptype.startswith("rule:"):
            rule_name = ptype[5:]
            fmt = rule_formats.get(rule_name, {})

            p = doc.add_paragraph()
            _format_paragraph(p, fmt.get("alignment", d_align), d_spacing)

            display = stripped.upper() if fmt.get("uppercase") else stripped
            run = p.add_run(display)
            _format_run(
                run, d_font,
                fmt.get("size", d_size),
                fmt.get("bold", d_bold),
                fmt.get("italic", d_italic),
                fmt.get("color", d_color),
                fmt.get("underline", False),
            )
            continue

        # --- DEFAULT BODY ---
        p = doc.add_paragraph()
        _format_paragraph(p, d_align, d_spacing)
        _add_text_with_inline_formatting(
            p, stripped, d_font, d_size, d_bold, d_italic, d_color, inline_rules
        )

    # Save
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


def _find_rule_format(keyword, paragraph_rules):
    """Find a paragraph rule by keyword in its name."""
    keyword = keyword.lower()
    for rule in paragraph_rules:
        if keyword in rule.get("name", "").lower():
            return rule.get("format", {})
    return None


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
