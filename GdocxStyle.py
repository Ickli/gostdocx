import copy
import json
import docx
from docx.styles.style import BaseStyle, ParagraphStyle, CharacterStyle
from docx.enum.text import WD_LINE_SPACING, WD_PARAGRAPH_ALIGNMENT, WD_COLOR_INDEX, WD_UNDERLINE
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt, Inches, Cm, RGBColor, Length
from docx import Document
from docx.text.run import Run, Font

'''
Guidelines for styles:
    For an example of style formats, look into styles/default.json.

    Recommendations:
    As the tool is created for GOST 7.32 document creation, please be aware
    of the following (list of things which surprised me):
    1. Centered text is truly centered only when first_line_indent = 0;
    2. Image caption is part of the same paragraph of which the image is
        (meaning that user must press SHIFT+ENTER to insert caption according
        to GOST) and should only be used right after the image;
    3. Therefore, style for image should contain description of image caption
        too.
    4. But I've added separate IMAGE_CAPTION style, so the tool is a bit more
        flexible.

    File and parsing description:
    1. Styles' file contains pairs of style names and style descriptions,
        and pairs of element names and element values;
    2. Every style description currently must contain "is_paragraph": true
        because the tool only supports ParagraphStyle settings;
    3. Feilds of style description directly correlate to attributes of
        ParagraphStyle, except:
            "font" field, which is object of format: {
                "name": "FONT_NAME",
                "size": SIZE_IN_PT,
                "color: [int, int, int]
            },
            "base_style" field whose value is used to subscript doc.styles,
            "alignment" field whose value is gotten via
                getattr(WD_PARAGRAPH_ALIGNMENT, ...),

    4. Every number is in Cm units, if not stated otherwise;
    5. Every element name is parsed as a special case in parse_raw_styles
        and is reflected in class Style as bunch of static attributes
'''

# names for fields of json, 
# these fields do not correlate to any ParagraphStyle attribute
FIELD_IS_PAR = "is_paragraph"
FIELD_UNORDERED_LIST_PREFIX = "unordered_list_prefix"
FIELD_IMAGE_CAPTION_PREFIX = "image_caption_prefix"
FIELD_IMAGE_CAPTION_INFIX = "image_caption_infix"

# You can change it
# Don't forget to add new default styles in set_defaults_if_not_set()
class Style:
    UNORDERED_LIST_PREFIX = None
    IMAGE_CAPTION_PREFIX = None
    IMAGE_CAPTION_INFIX = None

# Properties of this style are not dictated by any json
# All newly created styles will use these properties if not overriden
def create_default_par_style(style_name: str, doc: Document) -> ParagraphStyle:
    style = doc.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
    fmt = style.paragraph_format
    fmt.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    return style

def use_styles_from_file(filepath: str, doc: Document, to_override: bool = False):
    file = open(filepath, "r")
    json_string = file.read()
    parse_raw_styles(json_string, doc, to_override)

def parse_raw_styles(json_string: str, doc: Document, to_override: bool):
    raw_styles = json.loads(json_string)

    for style_name in raw_styles:
        if style_name == FIELD_UNORDERED_LIST_PREFIX:
            Style.UNORDERED_LIST_PREFIX = raw_styles[style_name]
            continue
        if style_name == FIELD_IMAGE_CAPTION_PREFIX:
            Style.IMAGE_CAPTION_PREFIX = raw_styles[style_name]
            continue
        if style_name == FIELD_IMAGE_CAPTION_INFIX:
            Style.IMAGE_CAPTION_INFIX = raw_styles[style_name]
            continue
        if style_name in doc.styles and not to_override:
            raise Exception(f"Style {style_name} encountered twice")
        raw_style = raw_styles[style_name]
        style = parse_raw_style(style_name, raw_style, doc)

def parse_raw_style(style_name: str, json_dict: dict[str, object], doc: Document) -> BaseStyle:
    if json_dict['is_paragraph']:
        del json_dict[FIELD_IS_PAR]
        style = parse_raw_par_style(style_name, json_dict, doc)
    else:
        del json_dict[FIELD_IS_PAR]
        style = parse_raw_char_style(style_name, json_dict, doc)
    return style

def parse_raw_par_style(style_name: str, json_dict: dict[str, object], doc: Document) -> ParagraphStyle:
    style = None
    try:
        style = doc.styles[style_name]
    except KeyError as e:
        style = create_default_par_style(style_name, doc)

    for name in json_dict:
        value = json_dict[name]
        if name == "font":
            parse_raw_font(style, value)
        elif name == "base_style":
            style.base_style = doc.styles[value]
        elif name == "alignment":
            style.paragraph_format.alignment = getattr(WD_PARAGRAPH_ALIGNMENT, value)
        elif name.endswith('indent'):
            setattr(style.paragraph_format, name, Cm(value))
        else:
            setattr(style.paragraph_format, name, value)
    return style

def parse_raw_char_style(style_name: str, json_dict: dict[str, object], doc: Document) -> BaseStyle:
    for name in json_dict:
        value = json_dict[name]
        if name == "font":
            parse_raw_font(style, value)
        else:
            raise Exception("Can't put in character style anything except 'font'")

def parse_raw_font(style: BaseStyle, font_dict: dict[str, object]):
    font = style.font
    for name in font_dict:
        value = font_dict[name]
        if name == "size":
            font.size = Pt(value)
        elif name == "color":
            font.color.rgb = RGBColor.from_string(value)
        elif name == "highlight_color":
            font.highlight_color = getattr(WD_COLOR_INDEX, value)
        elif name == "underline":
            if isinstance(value, bool):
                font.underline = value
            else:
                font.underline = getattr(WD_UNDERLINE, value)
        else:
            setattr(style.font, name, value)

DefaultStylesDoc = Document()

def init_default_styles(filepath: str):
    use_styles_from_file(filepath, DefaultStylesDoc)

# Slightly rewritten code from https://stackoverflow.com/questions/78733174/how-can-i-use-styles-from-an-existing-docx-file-in-my-new-document
#
# Assumes 'sname' style is in src and is not in dest
def copy_style(dest: Document, src: Document, sname: str):
    src_style = src.styles[sname]
    dest.styles.element.append(copy.copy(src_style.element))

def copy_styles(dest: Document, src: Document):
    src_stylenames = [st.name for st in src.styles]
    dest_stylenames = [st.name for st in dest.styles]

    for name in src_stylenames:
        if name in dest_stylenames:
            dest.styles[name].delete()
        copy_style(dest, src, name)

def use_default_styles(doc: Document):
    copy_styles(doc, DefaultStylesDoc)

###############################   Serialization   ##############################

def ser_par_align(style: ParagraphStyle) -> str:
    svars = vars(docx.enum.text.WD_PARAGRAPH_ALIGNMENT)

    for name in svars:
        if not name.startswith('_') and svars[name] == style.paragraph_format.alignment:
            return name
    return None

# I don't understand how paragraph styles' font correlate to run's font,
# so i serialize both.
def ser_font(font: Font) -> dict[str, any]:
    jdict = {}
    jdict['name'] = font.name
    if font.size is not None:
        jdict['size'] = font.size.pt
    if font.all_caps is not None:
        jdict['all_caps'] = font.all_caps
    if font.bold is not None:
        jdict['bold'] = font.bold
    color = font.color.rgb;
    if color is not None:
        jdict['color'] = str(color)
    if font.double_strike:
        jdict['double_strike'] = font.double_strike
    if font.highlight_color:
        svars = vars(WD_COLOR_INDEX)
        for name in svars:
            if svars[name] == font.highlight_color:
                jdict['highlight_color'] = name
                break
    if font.italic is not None:
        jdict['italic'] = font.italic
    if font.name is not None:
        jdict['name'] = font.name
    if font.strike is not None:
        jdict['strike'] = font.strike
    if font.subscript is not None:
        jdict['subscript'] = font.subscript
    if font.underline is not None:
        ul = font.underline
        if isinstance(ul, bool):
            jdict['underline'] = ul
        else:
            svars = vars(WD_UNDERLINE)
            for name in svars:
                if svars[name] == ul:
                    jdict['underline'] = name
                    break
    return jdict

def ser_line_spacing(style: ParagraphStyle) -> Length | None | float:
    lspacing = None
    pformat = style.paragraph_format

    if isinstance(pformat.line_spacing, Length):
        lspacing = pformat.line_spacing.cm
    elif isinstance(pformat.line_spacing, float):
        lspacing = pformat.line_spacing
    return lspacing

# doesn't include name of style
def ser_par_style(style: ParagraphStyle) -> dict[str, any]:
    jdict = {}
    pformat = style.paragraph_format

    jdict[FIELD_IS_PAR] = True
    if pformat.first_line_indent is not None:
        jdict['first_line_indent'] = pformat.first_line_indent.cm
    if pformat.left_indent is not None:
        jdict['left_indent'] = pformat.left_indent.cm
    if pformat.line_spacing is not None:
        jdict['line_spacing'] = ser_line_spacing(style)
    if pformat.space_before is not None:
        jdict['space_before'] = pformat.space_before.cm
    if pformat.space_after is not None:
        jdict['space_after'] = pformat.space_after.cm
    if style.base_style is not None:
        jdict['base_style'] = style.base_style.name
    try:
        jdict['alignment'] = ser_par_align(style)
    except:
        pass
    jdict['font'] = ser_font(style.font)

    return jdict

def ser_char_style(style: CharacterStyle, run: Run | None = None) -> dict[str, any]:
    jdict = {}
    jdict[FIELD_IS_PAR] = False
    if style.font is not None:
        jdict['font'] = ser_font(style.font)
    if run is not None:
        if not 'font' in jdict:
            jdict['font'] = {}
        fontdict = jdict['font']
        if 'italic' not in fontdict and run.italic is not None:
            fontdict['italic'] = run.italic
        if 'bold' not in fontdict and run.bold is not None:
            fontdict['bold'] = run.bold
    return jdict
