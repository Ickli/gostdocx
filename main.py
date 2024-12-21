import sys
import argparse
from docx import Document
from GdocxState import GdocxState
from typing import Type, Any
import GdocxParsing
import GdocxHandler
import GdocxStyle
import GdocxCommon
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement, ns

PATH_DEFAULT_STYLES = "styles/default.json"
SKIP_NUMBERING = False

# ! You can add something here !
# Will be added to GdocxState's registered_handlers
# key: string, value: function(state: GdocxState, macro_args: list[str]) -> handler
registered_macro_handlers: list[Type[Any]] = [
    GdocxHandler.EchoHandler
]

def process_with_current_handler(file, state: GdocxState):
    # this is for one-line macro
    if state.reached_macro_end:
        state.handler.finalize()
        state.reached_macro_end = False
        return

    line = file.readline()
    while(line != ""):
        new_handler = state.handle_or_get_new_handler(line)

        if new_handler is not None:
            # new_handler sees old_handler as state's current handler.
            # so it can perform some logic based on that
            old_handler = state.handler
            state.handler = new_handler

            state.indent += 1
            process_with_current_handler(file, state)

            state.indent -= 1
            state.handler = old_handler
        # this is for end macro
        if state.reached_macro_end:
            state.handler.finalize()
            state.reached_macro_end = False
            return

        line = file.readline()
    return

# Copied from https://stackoverflow.com/questions/56658872/add-page-number-using-python-docx
def create_element(name):
    return OxmlElement(name)

def create_attribute(element, name, value):
    element.set(ns.qn(name), value)


def add_page_number(run):
    fldChar1 = create_element('w:fldChar')
    create_attribute(fldChar1, 'w:fldCharType', 'begin')

    instrText = create_element('w:instrText')
    create_attribute(instrText, 'xml:space', 'preserve')
    instrText.text = "PAGE"

    fldChar2 = create_element('w:fldChar')
    create_attribute(fldChar2, 'w:fldCharType', 'end')

    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
# end copied

def add_footer_with_page_number(doc: Document):
    add_page_number(doc.sections[0].footer.paragraphs[0].add_run())
    doc.sections[0].footer.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

def process_file(filepath: str, filepath_out: str):
    file = open(filepath, "r")
    doc = Document()
    GdocxStyle.use_styles_from_file(PATH_DEFAULT_STYLES, doc)

    with GdocxState(doc, registered_macro_handlers) as state:
        state.strip_indent = GdocxParsing.STRIP_INDENT
        state.skip_empty = GdocxParsing.SKIP_EMPTY
        # here state is primary handler
        process_with_current_handler(file, state)

    file.close()

    if not SKIP_NUMBERING:
        add_footer_with_page_number(doc)
    doc.save(filepath_out)

def process_args() -> (str, str):
    inpath: str | None = None
    outpath: str | None = None

    prs = argparse.ArgumentParser(prog = "GostDocx",
        description = '''
Converts .txt files into .docx
>> To see an example, run 'python3 example.py'

Конвертирует .txt файлы в .docx
>> Чтобы увидеть пример, введи 'python3 example.py' ''',
        formatter_class = argparse.RawTextHelpFormatter
    )
    prs.add_argument('-i', '--input', help="Path to source file", type=str)
    prs.add_argument('-o', '--output', help="Path to output file", type=str)
    prs.add_argument('-s', '--strip-indent', help="strip indents of nested macros", action="store_true")
    prs.add_argument('-se', '--skip-empty', help="skip empty lines", action="store_true")
    prs.add_argument('-il', '--indent-length', help="Length of indent sequence", type=int)
    prs.add_argument('-ic', '--indent-char', help="Indent character", type=str)
    prs.add_argument('-n', '--skip-numbering', help="Don't put page number in footers of pages", action="store_true")

    args = prs.parse_args()
    inpath = args.input
    outpath = args.output
    il = args.indent_length
    ic = args.indent_char

    if ic == None:
        ic = GdocxParsing.INDENT_DEFAULT_CHAR
    if il == None:
        il = GdocxParsing.INDENT_DEFAULT_LENGTH
    GdocxParsing.INDENT_STRING = ic * il
    
    if inpath is None:
        print("ERROR: No input file is provided")
        prs.print_help()
        exit()
    if outpath is None:
        print("ERROR: No output file is provided")
        prs.print_help()
        exit()

    GdocxParsing.STRIP_INDENT = args.strip_indent
    GdocxParsing.SKIP_EMPTY = args.skip_empty
    SKIP_NUMBERING = args.skip_numbering

    return (inpath, outpath)

if __name__ == "__main__":
    inpath, outpath = process_args()
    process_file(inpath, outpath)
