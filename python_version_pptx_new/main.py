import math
import os
import sys
from tkinter.constants import CENTER

from openpyxl import Workbook, load_workbook
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn

SLIDE_WIDTH = Emu(7620000)
SLIDE_HEIGHT = Emu(19050000)

BLUE_DARK = RGBColor(0x00, 0x4A, 0xAD)
CIRCLE_FILL = RGBColor(0xBC, 0xDD, 0xFF)
CIRCLE_LINE = RGBColor(0x79, 0xA5, 0xDF)
TITLE_COLOR = RGBColor(0x1E, 0x2E, 0x5D)
SUBTITLE_COLOR = RGBColor(0x3A, 0x4F, 0x8B)
FOOTER_TEXT_COLOR = RGBColor(0xFF, 0xFF, 0xFF)
FOOTER_BG_COLOR = RGBColor(0x5F, 0x90, 0xD1)
BODY_COLOR = RGBColor(0x00, 0x00, 0x00)

presentation = Presentation()
presentation.slide_width = SLIDE_WIDTH
presentation.slide_height = SLIDE_HEIGHT


slide = None


FONT_NAME = "Segoe UI"

TITLE_SHIFT_DOWN = 150000


# White panel behind the title. Its underside is where the blocks start,
# so it doubles as the top edge of the first block's band.
TITLE_PANEL_TOP = 150068 + TITLE_SHIFT_DOWN
TITLE_PANEL_HEIGHT = 2265949
TITLE_PANEL_BOTTOM = TITLE_PANEL_TOP + TITLE_PANEL_HEIGHT


def set_font_name(run, name=FONT_NAME):
    """Sets latin, cs, and ea typefaces on a run.
    python-pptx's run.font.name only sets <a:latin>; Hebrew uses <a:cs> and
    East-Asian slots also inherit if unset, so we always set all three."""
    run.font.name = name
    run_properties = run._r.get_or_add_rPr()
    for tag in ('a:cs', 'a:ea'):
        typeface = run_properties.find(qn(tag))
        if typeface is None:
            typeface = run_properties.makeelement(qn(tag), {})
            run_properties.append(typeface)
        typeface.set('typeface', name)


def set_rtl(paragraph, align=PP_ALIGN.RIGHT):
    paragraph_properties = paragraph._p.get_or_add_pPr()
    paragraph_properties.set('rtl', '1')
    paragraph.alignment = align


def add_text(left, top, width, height, text, size_pt, color, bold=True, lang="he-IL",
             align=PP_ALIGN.RIGHT, auto_size=MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT,
             font=FONT_NAME, line_spacing=None):
    textbox = slide.shapes.add_textbox(Emu(left), Emu(top), Emu(width), Emu(height))
    text_frame = textbox.text_frame
    text_frame.word_wrap = True
    text_frame.auto_size = auto_size
    paragraph = text_frame.paragraphs[0]
    set_rtl(paragraph, align)
    run = paragraph.add_run()
    run.text = text
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.color.rgb = color
    set_font_name(run, font)
    set_lang(run, lang)
    if line_spacing is not None:
        paragraph.line_spacing = line_spacing
    return textbox


def set_lang(run, lang):
    run_properties = run._r.get_or_add_rPr()
    run_properties.set('lang', lang)


def add_multirun_text(left, top, width, height, runs,
                      align=PP_ALIGN.RIGHT, auto_size=MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT, line_spacing=None,
                      font=FONT_NAME):
    """runs: list of (text, color, size_pt, bold, lang) or (text, color, size_pt, bold, lang, font) tuples.
    Split mixed-font runs by adding an optional 6th element to override the function-level font."""
    textbox = slide.shapes.add_textbox(Emu(left), Emu(top), Emu(width), Emu(height))
    text_frame = textbox.text_frame
    text_frame.word_wrap = True
    text_frame.auto_size = auto_size
    paragraph = text_frame.paragraphs[0]
    set_rtl(paragraph, align)
    for run_data in runs:
        text, color, size_pt, bold, lang = run_data[:5]
        run_font = run_data[5] if len(run_data) > 5 else font
        run = paragraph.add_run()
        run.text = text
        run.font.size = Pt(size_pt)
        run.font.bold = bold
        run.font.color.rgb = color
        set_font_name(run, run_font)
        set_lang(run, lang)
    if line_spacing is not None:
        paragraph.line_spacing = line_spacing
    return textbox


def add_line(left, top, width, height, color=BLUE_DARK, weight_pt=5.0):
    connector = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Emu(left), Emu(top),
                                           Emu(left + width), Emu(top + height))
    connector.line.color.rgb = color
    connector.line.width = Pt(weight_pt)
    return connector


def set_reflection(shape, blur_radius=6350, start_alpha=50000, end_alpha=300,
                   end_pos=35000, distance=0):
    """Adds PowerPoint's reflection effect to a shape.
    python-pptx exposes no API for effects, so we append the <a:effectLst>
    element by hand. Defaults match the 'Tight Reflection, touching' preset;
    end_pos is how far down the fade runs (0-100000) and distance is the gap
    below the shape in EMU."""
    shape_properties = shape._element.spPr
    effect_list = shape_properties.find(qn('a:effectLst'))
    if effect_list is None:
        # Schema order puts effectLst after <a:ln>, which is already set.
        effect_list = shape_properties.makeelement(qn('a:effectLst'), {})
        shape_properties.append(effect_list)
    effect_list.append(effect_list.makeelement(qn('a:reflection'), {
        'blurRad': str(blur_radius),
        'stA': str(start_alpha),
        'stPos': '0',
        'endA': str(end_alpha),
        'endPos': str(end_pos),
        'dist': str(distance),
        'dir': '5400000',
        'fadeDir': '5400000',
        'sy': '-100000',
        'algn': 'bl',
        'rotWithShape': '0',
    }))
    return shape


def add_icon_bubble(left, top, size, number_text):
    circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, Emu(left), Emu(top), Emu(size), Emu(size))
    circle.fill.solid()
    circle.fill.fore_color.rgb = CIRCLE_FILL
    circle.line.color.rgb = CIRCLE_LINE
    circle.line.width = Pt(3.75)
    circle.shadow.inherit = False

    number_box = slide.shapes.add_textbox(Emu(left), Emu(top), Emu(size), Emu(size))
    text_frame = number_box.text_frame
    text_frame.word_wrap = False
    text_frame.vertical_anchor = 3  # MSO_ANCHOR.MIDDLE
    paragraph = text_frame.paragraphs[0]
    paragraph.alignment = PP_ALIGN.CENTER
    run = paragraph.add_run()
    run.text = number_text
    run.font.size = Pt(28)
    run.font.bold = True
    run.font.color.rgb = BLUE_DARK
    set_font_name(run, "Prachason Neue SemiExpanded Bold")
    set_lang(run, "he-IL")

    return slide.shapes.add_group_shape([circle, number_box])


def new_slide():
    global slide
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])  # blank layout

    background = slide.background
    background.fill.solid()
    background.fill.fore_color.rgb = RGBColor(0xF0, 0xF5, 0xFA)

    title_panel = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(129294), Emu(TITLE_PANEL_TOP), Emu(7361412), Emu(TITLE_PANEL_HEIGHT))
    title_panel.fill.solid()
    title_panel.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    title_panel.line.fill.background()

    slide.shapes.add_picture("PictureLogo.png", Emu(129294), Emu(161007 + TITLE_SHIFT_DOWN), Emu(1072462), Emu(1201987))
    slide.shapes.add_picture("PictureShield.png", Emu(6761525), Emu(312339 + TITLE_SHIFT_DOWN), Emu(638880), Emu(899321))

    footer_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(-134306), Emu(18432419), Emu(7888611), Emu(757937))
    footer_bar.fill.solid()
    footer_bar.fill.fore_color.rgb = FOOTER_BG_COLOR
    footer_bar.line.fill.background()

    add_multirun_text(712460, 353130 + TITLE_SHIFT_DOWN, 6096000, 820738, [
        ("מה חדש ב", TITLE_COLOR, 64, True, "he-IL", "Choco Bold"),
        ("SAP", TITLE_COLOR, 64, True, "en-US", "Cascadia Mono SemiBold"),
        ("?", TITLE_COLOR, 64, True, "ar-EG", "Choco Bold"),
    ], align=PP_ALIGN.CENTER, auto_size=MSO_AUTO_SIZE.NONE, line_spacing=Pt(64))
    add_multirun_text(683920, 1329693 + TITLE_SHIFT_DOWN, 6096000, 589905, [
        ("גרסת אוגוסט ", SUBTITLE_COLOR, 23, True, "he-IL"),
        ("2026", SUBTITLE_COLOR, 23, True, "en-US"),
        (" - מועד עליה 16/9/26", SUBTITLE_COLOR, 23, True, "he-IL"),
        (" ", SUBTITLE_COLOR, 23, True, "ar-EG"),
        ('\nמדור מע\u201dמ', RGBColor(0x68, 0xBB, 0xE3), 23, True, "he-IL"),
    ], align=PP_ALIGN.CENTER, auto_size=MSO_AUTO_SIZE.NONE, line_spacing=Pt(23))
    add_text(684575,1989975 + TITLE_SHIFT_DOWN, 6096000, 261867, "גרסאות עתידיות 21/10/26 + 7/10/26", 14, SUBTITLE_COLOR,
             align=PP_ALIGN.CENTER, auto_size=MSO_AUTO_SIZE.NONE)
    add_text(0, 18570540, SLIDE_WIDTH, 296235, "כשטכנולוגיה פוגשת אנשים", 17, FOOTER_TEXT_COLOR,
             True, "he-IL", align=PP_ALIGN.CENTER, auto_size=MSO_AUTO_SIZE.NONE,
             font="Guttman Yad-Brush")


# The header reads 'client | jira' in four tones of blue, one run each.
HEADER_CLIENT_COLOR = RGBColor(0x00, 0x4A, 0xAD)
HEADER_SEPARATOR_COLOR = RGBColor(0x24, 0x79, 0xDF)
HEADER_SPACER_COLOR = RGBColor(0x38, 0xB6, 0xFF)
HEADER_JIRA_COLOR = RGBColor(0x79, 0xA5, 0xDF)

HEADER_SIZE_PT = 18
# The body size is the only thing that adapts to content; see the fitting
# section below for what it is fitted against.
BODY_SIZE_MIN_PT = 16
BODY_SIZE_MAX_PT = 16
BODY_SIZE_STEP_PT = 0.5
# Body paragraphs only - the header stays at single spacing.
BODY_LINE_SPACING = 1.5
BUBBLE_SIZE = 856239


# Per-slot layout (1-8), in EMU: header/body as (left, width). Their tops come
# from the block's band and their heights from the text itself, so only the
# connector lines carry absolute positions; the bubble is centred on the
# vertical one.
_BLOCK_GEOMETRY = {
    1: dict(header=(218909, 5647223),
            body=(539083, 5143320),
            lines=[(6323881, 3438646, 0, 748753), (1060308, 4163587, 5282623, 0)]),
    2: dict(header=(1617765, 5647223),
            body=(1521457, 5559802),
            lines=[(1081700, 4139775, 0, 1956226), (1060308, 6096001, 5282623, 0)]),
    3: dict(header=(164917, 5647223),
            body=(571680, 5143320),
            lines=[(6323881, 6072188, 0, 1954800), (1055852, 8005218, 5282623, 0)]),
    4: dict(header=(1377359, 5883173),
            body=(1521457, 5559802),
            lines=[(1077244, 7981406, 0, 1956226), (1055852, 9937633, 5282623, 0)]),
    5: dict(header=(213537, 5647223),
            body=(304800, 5372231),
            lines=[(6319425, 9913820, 0, 1954800), (1055852, 11868620, 5282623, 0)]),
    6: dict(header=(1613309, 5647223),
            body=(1676400, 5559802),
            lines=[(1077244, 11844808, 0, 1956226), (1055851, 13799607, 5282623, 0)]),
    7: dict(header=(224975, 5647223),
            body=(304800, 5383669),
            lines=[(6319424, 13775794, 0, 1954800), (1055850, 15737851, 5282623, 0)]),
    8: dict(header=(1613309, 5647223),
            body=(1457631, 5559802),
            lines=[(1077242, 15714039, 0, 1956226), (1055850, 17670265, 5282623, 0)]),
}


BLOCKS_PER_SLIDE = len(_BLOCK_GEOMETRY)  # 8 - a full page of block slots



EMU_PER_PT = 12700
# python-pptx textbox defaults: 0.1" left/right, 0.05" top/bottom insets.
_INSET_X_EMU = 2 * 91440
_INSET_Y_EMU = 2 * 45720

# Content sits at the top of its band rather than centred in it: this is the
# gap between the band's top edge and the header, and what is left over drops
# out as slack at the bottom.
BLOCK_TOP_PADDING_PT = 16
BLOCK_TARGET_MARGIN_PT = 30  # preferred slack below the body
_BLOCK_MIN_MARGIN_PT = 4

HEADER_TO_BODY_OFFSET_EMU = 391285

_MEASURE_PX = 100  # measure once at this size; widths scale linearly with pt
_FALLBACK_CHAR_WIDTH_RATIO = 0.5  # mean glyph width if the font can't be read
_FALLBACK_LINE_RATIO = 1.33       # Segoe UI ascent+descent per em


_FONT_CACHE = {}


def _measure_font(bold=False):
    if bold not in _FONT_CACHE:
        file_stem = FONT_NAME.replace(" ", "").lower()
        try:
            from PIL import ImageFont  # ships with python-pptx
            _FONT_CACHE[bold] = ImageFont.truetype(f"{file_stem}{'b' if bold else ''}.ttf",
                                                   _MEASURE_PX)
        except Exception:
            _FONT_CACHE[bold] = None
    return _FONT_CACHE[bold]


def _line_height_ratio():
    """Single-line spacing as a multiple of font size, taken from the font's
    own ascent+descent - the same metric PowerPoint lays lines out with."""
    font = _measure_font()
    if font is None:
        return _FALLBACK_LINE_RATIO
    ascent, descent = font.getmetrics()
    return (ascent + descent) / _MEASURE_PX


def _text_width_pt(text, size_pt, bold=False):
    font = _measure_font(bold)
    if font is None:
        return _FALLBACK_CHAR_WIDTH_RATIO * size_pt * len(text)
    return font.getlength(text) * size_pt / _MEASURE_PX


def _wrapped_line_count(text, size_pt, width_pt, bold=False):
    """Greedy word wrap - the same break points PowerPoint picks, since
    both break on whitespace at the first word that overruns the line."""
    lines, line_so_far = 1, ""
    for word in text.split():
        word_width = _text_width_pt(word, size_pt, bold)
        if word_width > width_pt:
            # Wider than the box on its own: PowerPoint breaks it mid-word
            # rather than letting it overrun, so it spans several lines.
            if line_so_far:
                lines += 1
            lines += math.ceil(word_width / width_pt) - 1
            line_so_far = ""
            continue
        extended = f"{line_so_far} {word}" if line_so_far else word
        if not line_so_far or _text_width_pt(extended, size_pt, bold) <= width_pt:
            line_so_far = extended
        else:
            lines += 1
            line_so_far = word
    return lines


def _candidate_sizes():
    size = BODY_SIZE_MIN_PT
    while size <= BODY_SIZE_MAX_PT:
        yield size
        size += BODY_SIZE_STEP_PT


def _header_text(client, jira):
    """The ' | ' separator is added here, never typed in Excel, so that
    measuring and drawing can't drift apart."""
    return f"{client} | {jira}"


def _block_band(position):
    """(top, bottom) of the slide band a block owns, in EMU.

    Bounded by the connector lines separating the blocks; the first block
    on a page starts under the white title panel instead."""
    top = (TITLE_PANEL_BOTTOM if position == 1
           else _BLOCK_GEOMETRY[position - 1]['lines'][1][1])
    return top, _BLOCK_GEOMETRY[position]['lines'][1][1]


def _header_line_count(header_width_emu, client, jira):
    width_pt = (header_width_emu - _INSET_X_EMU) / EMU_PER_PT
    return _wrapped_line_count(_header_text(client, jira), HEADER_SIZE_PT, width_pt, bold=True)


def _header_height_emu(line_count):
    return round(line_count * HEADER_SIZE_PT * _line_height_ratio() * EMU_PER_PT) + _INSET_Y_EMU


def _header_to_body_emu(line_count):
    """Header top to body top: the fixed offset, plus room for any extra
    line a long header wraps onto, so the body never rides up into it."""
    wrapped_extra_emu = (line_count - 1) * HEADER_SIZE_PT * _line_height_ratio() * EMU_PER_PT
    return HEADER_TO_BODY_OFFSET_EMU + round(wrapped_extra_emu)


def _body_height_emu(text, size_pt, box_width_emu):
    width_pt = (box_width_emu - _INSET_X_EMU) / EMU_PER_PT
    lines = _wrapped_line_count(text, size_pt, width_pt)
    # PowerPoint scales every line by the spacing multiple, the first included.
    line_height = size_pt * _line_height_ratio() * BODY_LINE_SPACING
    return round(lines * line_height * EMU_PER_PT) + _INSET_Y_EMU


def _fit_body_size(text, box_width_emu, header_offset_emu, band_height_emu):
    """Body size whose leftover margin lands nearest BLOCK_TARGET_MARGIN_PT.
    Ties go to the larger size; falls back to the minimum if nothing fits.
    The margin measured is the slack below the body, since the top gap is
    fixed at BLOCK_TOP_PADDING_PT."""
    best_score, best_size = None, BODY_SIZE_MIN_PT
    top_padding = BLOCK_TOP_PADDING_PT * EMU_PER_PT
    for size in _candidate_sizes():
        content_height = header_offset_emu + _body_height_emu(text, size, box_width_emu)
        margin_pt = (band_height_emu - top_padding - content_height) / EMU_PER_PT
        if margin_pt < _BLOCK_MIN_MARGIN_PT:
            continue
        score = (abs(margin_pt - BLOCK_TARGET_MARGIN_PT), -size)
        if best_score is None or score < best_score:
            best_score, best_size = score, size
    return best_size


def _draw_version_note_block(position, client, jira, body_text, number):
    layout = _BLOCK_GEOMETRY[position]

    header_runs = [(client + ' ', HEADER_CLIENT_COLOR, HEADER_SIZE_PT, True, "he-IL"),
                   ('|', HEADER_SEPARATOR_COLOR, HEADER_SIZE_PT, True, "ar-EG"),
                   (' ', HEADER_SPACER_COLOR, HEADER_SIZE_PT, True, "ar-EG"),
                   (jira, HEADER_JIRA_COLOR, HEADER_SIZE_PT, True, "he-IL")]

    header_left, header_width = layout['header']
    body_left, body_width = layout['body']

    band_top, band_bottom = _block_band(position)
    band_height = band_bottom - band_top
    header_lines = _header_line_count(header_width, client, jira)
    header_offset = _header_to_body_emu(header_lines)
    body_size = _fit_body_size(body_text, body_width, header_offset, band_height)

    # Both boxes are auto-fit, so their declared height is just the measured
    # height of the text they hold.
    header_height = _header_height_emu(header_lines)
    body_height = _body_height_emu(body_text, body_size, body_width)

    header_top = band_top + BLOCK_TOP_PADDING_PT * EMU_PER_PT
    body_top = header_top + header_offset

    add_multirun_text(header_left, header_top, header_width, header_height, header_runs)
    add_text(body_left, body_top, body_width, body_height, body_text, body_size,
             BODY_COLOR, bold=False, line_spacing=BODY_LINE_SPACING)

    vertical_line, horizontal_line = layout['lines']
    add_line(*vertical_line)
    add_line(*horizontal_line)

    # Centred on the vertical connector line, both across and along it.
    line_left, line_top, _, line_height = vertical_line
    bubble_left = line_left - BUBBLE_SIZE // 2
    bubble_top = line_top + (line_height - BUBBLE_SIZE) // 2
    add_icon_bubble(bubble_left, bubble_top, BUBBLE_SIZE, str(number))


def render_version_note(blocks):
    for first_index in range(0, max(len(blocks), 1), BLOCKS_PER_SLIDE):
        page_blocks = blocks[first_index:first_index + BLOCKS_PER_SLIDE]
        new_slide()
        for slot, (client, jira, body_text) in enumerate(page_blocks, start=1):
            _draw_version_note_block(slot, client, jira, body_text, first_index + slot)
        if len(page_blocks) == BLOCKS_PER_SLIDE:
            set_reflection(add_line(6319423, 17646452, 0, 468000))
        slide.shapes.add_picture("PicturePeople.png", Emu(28789), Emu(17368014), Emu(2129096), Emu(1681986))




VERSION_NOTE_XLSX = sys.argv[1] if len(sys.argv) > 1 else "version_note.xlsx"

COLUMN_TITLES = ("שם ג'ירה", "לקוח", "פירוט")
# Any cell matching one of these marks a row as headings rather than data.
_HEADING_CELLS = {title.lower() for title in COLUMN_TITLES}


def _is_heading_row(values):
    return any(str(value).strip().lower() in _HEADING_CELLS
               for value in values if value is not None)


def load_version_note_from_excel(path):
    """Reads the workbook into a list of (client, jira, body_text) tuples,
    skipping the heading row and any blank ones."""
    sheet = load_workbook(path, data_only=True).worksheets[0]
    blocks = []
    for row in sheet.iter_rows(min_col=1, max_col=3, values_only=True):
        jira, client, body = (str(value or "").strip() for value in row)
        if _is_heading_row(row):
            continue
        if not client and not jira and not body:
            continue
        blocks.append((client, jira, body))
    return blocks


def write_version_note_template(path):
    """Writes an empty, correctly shaped workbook to fill in by hand."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "VersionNote"
    sheet.sheet_view.rightToLeft = True  # the content is Hebrew
    sheet.append(list(COLUMN_TITLES))
    for column, width in zip("ABC", (45, 22, 90)):
        sheet.column_dimensions[column].width = width
    workbook.save(path)


def load_version_note(path=VERSION_NOTE_XLSX):
    """The workbook is the only source of entries. If it isn't there yet,
    lay down a blank one to fill in and stop - there is nothing to render."""
    if not os.path.exists(path):
        write_version_note_template(path)
        print(f"Created empty {path} - fill it in and re-run")
        sys.exit(0)
    blocks = load_version_note_from_excel(path)
    print(f"Loaded {len(blocks)} version note entries from {path}")
    return blocks


# Blocks 9-16 land on a second slide, 17-24 on a third, and so on.
render_version_note(load_version_note())


OUTPUT_PPTX = "output.pptx"
FALLBACK_PPTX = "output_new.pptx"
try:
    presentation.save(OUTPUT_PPTX)
    print(f"Saved: {OUTPUT_PPTX}")
except PermissionError:
    presentation.save(FALLBACK_PPTX)
    print(f"{OUTPUT_PPTX} is open — saved as {FALLBACK_PPTX}")