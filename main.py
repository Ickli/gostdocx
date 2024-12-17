import sys
from docx import Document
from GdocxState import GdocxState
from typing import Callable
import GdocxParsing
import GdocxHandler

# handler trait:
#   # Called on every occurrence of custom macro in file
#   __init__(self, state: GdocxState, macro_args: list[str]);
#
#   # Called on every line inside custom macro
#   process_line(self, line: str, info: GdocxParsing.LineInfo);
#
#   # Called when reached end of scope of custom macro
#   finalize(self);

# ! You can add something here !
# Will be added to GdocxState's registered_handlers
# key: string, value: function(state: GdocxState, macro_args: list[str]) -> handler
registered_macro_handlers: dict[str, Callable[[GdocxState, list[str]], object]] = {
    "echo": lambda state, macro_args: GdocxHandler.EchoHandler(state, macro_args)
}

def gdocx_delegate_to_handler(file, handler, state: GdocxState):
    if state.reached_macro_end:
        handler.finalize()
        state.reached_macro_end = False
        return

    line = file.readline()
    while(line != ""):
        line, info = GdocxParsing.process_line(line)
        new_handler = state.handle_or_get_new_handler(line, info, handler)

        if new_handler is not None:
            gdocx_delegate_to_handler(file, new_handler, state)
        if state.reached_macro_end:
            handler.finalize()
            state.reached_macro_end = False
            return

        line = file.readline()
    return

def gdocx_process_file(filepath: str, filepath_out: str) -> str:
    file = open(filepath, "r")
    doc = Document()

    with GdocxState(doc, registered_macro_handlers) as state:
        # here state is primary handler
        gdocx_delegate_to_handler(file, state, state)

    file.close()
    doc.save(filepath_out)

if __name__ == "__main__":
    for i in range(1, len(sys.argv)):
        gdocx_process_file(sys.argv[i], "test.docx")
