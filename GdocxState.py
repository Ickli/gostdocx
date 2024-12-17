from typing import Callable
from docx import Document
import GdocxHandler
import GdocxParsing
import GdocxStyle

default_handlers: dict[str, Callable[['GdocxState', list[str]], object]] = {
        "unordered-list-item": lambda state, macro_args: GdocxHandler.UnorderedListHandler(state, macro_args),
        "load-style": lambda state, macro_args: GdocxHandler.LoadStyleHandler(state, macro_args),
}

# Document passed to ctor must outlive GdocxState
class GdocxState:
    def __init__(self, 
        doc: Document, 
        handlers: dict[str, Callable[['GdocxState', list[str]], object]]
    ):
        self.doc = doc
        self.cur_paragraph_lines = []
        self.current_style = doc.styles['Normal']
        self.registered_handlers = default_handlers | handlers
        self.reached_macro_end = False
        GdocxStyle.set_defaults_if_not_set(doc)

    # Returns new handler, if macro is encountered;
    # otherwise, returns None
    def handle_or_get_new_handler(self,
        line: str, 
        info: GdocxParsing.LineInfo,
        handler) -> object | None:

        if info.type == GdocxParsing.INFO_TYPE_MACRO:
            return self.process_macro_line(line, info)

        handler.process_line(line, info)
        return None

    def process_line(self, line: str, info: GdocxParsing.LineInfo):
        # GdocxParsing.INFO_TYPE_MACRO is handled in caller 'handle_or_get_new_handler'
        if info.type == GdocxParsing.INFO_TYPE_PLAIN_LINE:
            self.process_plain_line(line, info)
        elif info.type == GdocxParsing.INFO_TYPE_HEADER:
            self.process_header(line, info)

    # Returns new handler or None
    def process_macro_line(self, 
        line: str, 
        info: GdocxParsing.LineInfo
    ) -> object | None:
        macro_type = GdocxParsing.get_macro_type(info.line_stripped)
        new_handler = None

        if macro_type == GdocxParsing.MACRO_TYPE_START or macro_type == GdocxParsing.MACRO_TYPE_ONE_LINE:
            args = GdocxParsing.parse_macro_args(info.line_stripped)
            if len(args) == 0:
                raise Exception("Empty macro")
            macro_name = args[0]
            if macro_name not in self.registered_handlers:
                raise Exception("Couldn't find macro: %s" % macro_name)
            else:
                new_handler = self.registered_handlers[macro_name](self, args[1:])

            self.reached_macro_end = (macro_type == GdocxParsing.MACRO_TYPE_END or macro_type == GdocxParsing.MACRO_TYPE_ONE_LINE)

        return new_handler
    
    def process_plain_line(self, line: str, info: GdocxParsing.LineInfo):
        self.cur_paragraph_lines.append(line)

    def process_header(self, line: str, info: GdocxParsing.LineInfo):
        self.doc.add_heading(GdocxParsing.get_header_string(line), 0)

    def write_paragraph(self):
        par = '\n'.join(self.cur_paragraph_lines)
        self.doc.add_paragraph(par)

    def flush_paragraph(self):
        if len(self.cur_paragraph_lines) != 0:
            self.write_paragraph()
            self.cur_paragraph_lines = []

    def finalize(self):
        self.flush_paragraph()

    def __enter__(self):
        return self
    
    def __exit__(self, exc_type, exc_value, traceback):
        self.finalize()

    def echo(line: str):
        print("echo:", line)
