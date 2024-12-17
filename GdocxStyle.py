import json
from docx.styles.style import BaseStyle, ParagraphStyle
from docx.enum.text import WD_LINE_SPACING
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt
from docx import Document

# names for fields of json, 
# these fields do not correlate to any docx.Style attribute
FIELD_IS_PAR = "is_paragraph"

# Properties of this style are not dictated by any json
# All newly created styles will use these properties if not overriden
def create_default_par_style(style_name: str, doc: Document) -> ParagraphStyle:
    style = doc.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
    fmt = style.paragraph_format
    fmt.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    return style

class Style:
    UNORDERED_LIST = None
    ORDERED_LIST = None

    # Names for default styles
    DNAME_UNORDERED_LIST = "List Bullet"
    DNAME_ORDERED_LIST = "List Number"

def set_defaults_if_not_set(doc: Document):
    if Style.UNORDERED_LIST is None:
        Style.UNORDERED_LIST = doc.styles[Style.DNAME_UNORDERED_LIST]
    if Style.ORDERED_LIST is None:
        Style.ORDERED_LIST = doc.styles[Style.DNAME_ORDERED_LIST]


def parse_raw_styles(json_string: str, doc: Document):
    raw_styles = json.loads(json_string)

    for style_name in raw_styles:
        raw_style = raw_styles[style_name]
        if style_name in doc.styles:
            raise Exception(f"Style {style_name} encountered twice")
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
        elif name.endswith('indent'):
            setattr(style.paragraph_format, name, Pt(value))
        else:
            setattr(style.paragraph_format, name, value)

def parse_raw_font(style: BaseStyle, font_dict: dict[str, object]):
    for name, value in font_dict:
        if name == "size":
            style.font.size = Pt(value)
        else:
            setattr(style.font, name, value)
