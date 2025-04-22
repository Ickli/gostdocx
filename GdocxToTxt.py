from docx import Document
from pathlib import Path
from docx.text.run import Run
from docx.text.paragraph import Paragraph
from docx.styles.style import BaseStyle, ParagraphStyle, CharacterStyle
from docx.text.pagebreak import RenderedPageBreak
import GdocxStyle
import docx2txt
import json
import os

USE_DEFAULT_STYLES = False

DSTYLE_PARAGRAPH = 0
DSTYLE_HEADING1 = 1
DSTYLE_HEADING2 = 2

DSTYLES = {
    DSTYLE_PARAGRAPH: "paragraph",
    DSTYLE_HEADING1: "Heading 1",
    DSTYLE_HEADING2: "Heading 2",
}

def get_default_style(style: BaseStyle) -> int:
    if style.name == "Normal" or style.base_style.name == "Normal":
        return DSTYLE_PARAGRAPH
    if style.name == "Heading 1" or style.base_style.name == "Heading 1":
        return DSTYLE_HEADING1
    if style.name == "Heading 2" or style.base_style.name == "Heading 2":
        return DSTYLE_HEADING2
    return DSTYLE_PARAGRAPH

_FreeStyleIndex = 0
# returns hash of style, contained in style_hashes
# dict contains: keys = hashes of styles, values = pair of style name and style serialized into dict
def get_or_create_style(style: BaseStyle, style_hashes: dict[int, (str, dict[str, any])], run: Run | None = None) -> int:
    if USE_DEFAULT_STYLES:
        return get_default_style(style)
    else:
        return get_or_create_style_(style, style_hashes, run)

def get_or_create_style_(style: BaseStyle, style_hashes: dict[int, (str, dict[str, any])], run: Run | None = None) -> int:
    global _FreeStyleIndex

    jsonstr = ""
    jsondict = {}
    if isinstance(style, ParagraphStyle):
        jsondict = GdocxStyle.ser_par_style(style);
    else:
        jsondict = GdocxStyle.ser_char_style(style, run);
    jsonstr = json.dumps(jsondict)

    jsonhash = hash(jsonstr)
    if jsonhash in style_hashes:
        return jsonhash
    
    stylename = f"style{_FreeStyleIndex}"
    _FreeStyleIndex += 1
    style_hashes[jsonhash] = (stylename, jsondict)
    return jsonhash

def compose_style_hashes(style_hashes: dict[int, (str, dict[str, any])]) -> dict[str, any]:
    outdict = {}
    for stylehash in style_hashes:
        name_dict_pair = style_hashes[stylehash]
        outdict[name_dict_pair[0]] = name_dict_pair[1]
    return outdict

def macro_str_from_run(run: Run, run_text: str, style_hashes: dict[int, (str, str)]) -> str:

    stylename = "paragraph"
    if run.style is not None:
        stylehash = get_or_create_style(run.style, style_hashes, run)
        stylename = style_hashes[stylehash][0]
    return f'''(run-styled {stylename}
    {run_text}
)
'''

def macro_str_par_open(par: Paragraph, style_hashes: dict[int, (str, dict[str, any])]) -> str:
    stylehash = get_or_create_style(par.style, style_hashes)
    stylename = style_hashes[stylehash][0]
    return f'''(paragraph-styled {stylename}
'''

def macro_str_par_close() -> str:
    return ')\n'

def macro_str_from_image_path(path: str) -> str:
    # TODO
    return f"(img {_qt}{path}{_qt}){_nl}"

def macro_str_page_break() -> str:
    return "(page-break)\n"

def extract_images(docpath: str, dirpath: str):
    docx2txt.process(docpath, dirpath)

def docx_to_txt(docpath: str, outfilename: str, outdirpath: str):
    style_hashes = {}
    outdirpath = Path(outdirpath)

    outdirpath.mkdir(parents=True, exist_ok=True)
    extract_images(docpath, outdirpath)

    outstr = ""

    doc = Document(docpath)
    for s in (doc.styles.latent_styles):
        print(vars(s))
    doc_contents = doc.iter_inner_content()

    for doc_elem in doc_contents:
        if isinstance(doc_elem, Paragraph):
            par = doc_elem
            outstr += macro_str_par_open(doc_elem, style_hashes)

            if doc_elem.style.paragraph_format.page_break_before:
                outstr += macro_str_page_break()
            
            for run in par.iter_inner_content():
                if not isinstance(run, Run):
                    continue

                for run_elem in run.iter_inner_content():
                    if isinstance(run_elem, RenderedPageBreak):
                        pass # python-docx doesn't see page breaks
                    elif isinstance(run_elem, str):
                        outstr += macro_str_from_run(run, run_elem, style_hashes)
                    else: # Drawing
                        print(vars(run_elem._element))
                        img_name = run_elem._element._inline.graphic.graphicData.pic.nvPicPr.cNvPr.name
                        print("about to acess image:", img_name, outdirpath / img_name)
                        outstr += macro_str_from_image_path(outdirpath/img_name)

            outstr += macro_str_par_close()
        else: # doc_elem is Table
            pass

    print(outstr)
    print(json.dumps(compose_style_hashes(style_hashes), indent=4))
    # TODO
#     with open(outdirpath/outfilename, 'w') as file:
#         file.write(outstr)
