import sys
import argparse
from docx import Document
from GdocxState import GdocxState
from typing import Type, Any
import GdocxParsing
import GdocxHandler
import GdocxStyle
import GdocxCommon

PATH_DEFAULT_STYLES = "styles/default.json"

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

def process_file(filepath: str, filepath_out: str):
    file = open(filepath, "r")
    doc = Document()
    GdocxStyle.use_styles_from_file(PATH_DEFAULT_STYLES, doc)

    with GdocxState(doc, registered_macro_handlers) as state:
        state.strip_indent = GdocxParsing.STRIP_INDENT
        # here state is primary handler
        process_with_current_handler(file, state)

    file.close()
    doc.save(filepath_out)
    for warn in GdocxCommon.Warnings:
        print(warn)

def process_args() -> (str, str):
    inpath: str | None = None
    outpath: str | None = None

    prs = argparse.ArgumentParser(prog = "GostDocx",
        description = "Converts .txt files into .docx"
    )
    prs.add_argument('-i', '--input', help="Path to source file", type=str)
    prs.add_argument('-o', '--output', help="Path to output file", type=str)
    prs.add_argument('-s', '--strip-indent', help="strip indents of nested macros", action="store_true")
    prs.add_argument('-il', '--indent-length', help="Length of indent sequence", type=int)
    prs.add_argument('-ic', '--indent-char', help="Indent character", type=str)

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
    return (inpath, outpath)

if __name__ == "__main__":
    inpath, outpath = process_args()
    process_file(inpath, outpath)
