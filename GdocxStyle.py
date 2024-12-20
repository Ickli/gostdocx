import json
from docx.styles.style import BaseStyle, ParagraphStyle
from docx.enum.text import WD_LINE_SPACING, WD_PARAGRAPH_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt, Inches, Cm
from docx import Document

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
                "size": SIZE_IN_PT
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

    PARAGRAPH = None
    UNORDERED_LIST = None
    ORDERED_LIST = None
    IMAGE_PARAGRAPH = None
    HEADING_1 = None
    HEADING_2 = None
    # Note: this style is not used in GOST 7.32 really;
    #   instead, description of caption's style is 
    #   dictated by IMAGE_PARAGRAPH.
    # But you still can set the flag in ImageCaptionHandler,
    #   to use this style
    IMAGE_CAPTION = None

    # Names for default styles
    DNAME_PARAGRAPH = "paragraph"
    DNAME_UNORDERED_LIST = "list"
    DNAME_ORDERED_LIST = "list"
    DNAME_IMAGE_PARAGRAPH = "image-paragraph"
    DNAME_IMAGE_CAPTION = "image-caption"
    DNAME_HEADING_1 = "heading-1"
    DNAME_HEADING_2 = "heading-2"
    DNAME_HEADING_1_NON_TOC = "heading-1-non-toc"
    DNAME_HEADING_2_NON_TOC = "heading-2-non-toc"

def set_defaults_if_not_set(doc: Document):
    if Style.UNORDERED_LIST is None:
        Style.UNORDERED_LIST = doc.styles[Style.DNAME_UNORDERED_LIST]
    if Style.ORDERED_LIST is None:
        Style.ORDERED_LIST = doc.styles[Style.DNAME_ORDERED_LIST]
    if Style.IMAGE_PARAGRAPH is None:
        Style.IMAGE_PARAGRAPH = doc.styles[Style.DNAME_IMAGE_PARAGRAPH]
    if Style.PARAGRAPH is None:
        Style.PARAGRAPH = doc.styles[Style.DNAME_PARAGRAPH]

# Properties of this style are not dictated by any json
# All newly created styles will use these properties if not overriden
def create_default_par_style(style_name: str, doc: Document) -> ParagraphStyle:
    style = doc.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
    fmt = style.paragraph_format
    fmt.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    return style

def use_styles_from_file(filepath: str, doc: Document):
    file = open(filepath, "r")
    json_string = file.read()
    parse_raw_styles(json_string, doc)

def parse_raw_styles(json_string: str, doc: Document):
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
        if style_name in doc.styles:
            raise Exception(f"Style {style_name} encountered twice")
        raw_style = raw_styles[style_name]
        style = parse_raw_style(style_name, raw_style, doc)

def parse_raw_style(style_name: str, json_dict: dict[str, object], doc: Document) -> BaseStyle:
    if FIELD_IS_PAR not in json_dict or not json_dict[FIELD_IS_PAR]:
        raise Exception(f"Not implemented: style {style_name}, can't process styles whose is_paragraph is not set to true")

    style = create_default_par_style(style_name, doc)
    del json_dict[FIELD_IS_PAR]

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

def parse_raw_font(style: BaseStyle, font_dict: dict[str, object]):
    for name in font_dict:
        value = font_dict[name]
        if name == "size":
            style.font.size = Pt(value)
        else:
            setattr(style.font, name, value)
