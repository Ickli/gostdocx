import GdocxParsing
import GdocxStyle

# handler trait:
#   Name: str
#
#   # Called on every occurrence of custom macro in file
#   __init__(self, state: GdocxState, macro_args: list[str]);
#
#   # Called on every line inside custom macro
#   process_line(self, line: str, info: GdocxParsing.LineInfo);
#
#   # Called when reached end of scope of custom macro
#   finalize(self);

class EchoHandler:
    Name = "echo"

    def __init__(self, state: 'GdocxState', macro_args: list[str]):
        print('echo: ', *macro_args)
        pass

    def process_line(self, line: str, info: GdocxParsing.LineInfo):
        raise Exception("You must not place content inside echo")

    def finalize(self):
        pass

# Converts paragraphs to unordered list items
class UnorderedListItemHandler:
    Name = "unordered-list-item"

    def __init__(self, state: 'GdocxState', macro_args: list[str]):
        self.state = state
        self.cur_paragraph_lines = []
        self.is_first_line_processed = False

    def process_line(self, line: str, info: GdocxParsing.LineInfo):
        if not self.is_first_line_processed:
            line_stripped = GdocxStyle.Style.UNORDERED_LIST_PREFIX + info.line_stripped
            self.is_first_line_processed = True
        self.cur_paragraph_lines.append(line_stripped)

    def finalize(self):
        par_content = '\n'.join(self.cur_paragraph_lines)
        par = self.state.doc.add_paragraph(par_content)
        par.style = GdocxStyle.Style.UNORDERED_LIST

# Applies style to macro contents, treating it as a paragraph
class ParStyleHandler:
    Name = "paragraph-styled"

    def __init__(self, state: 'GdocxState', macro_args: list[str]):
        if len(macro_args) == 0:
            raise Exception("Style macro must contain style name")
        self.state = state
        self.style_name = macro_args[0]
        self.cur_paragraph_lines = []

    def process_line(self, line: str, info: GdocxParsing.LineInfo):
        self.cur_paragraph_lines.append(info.line_stripped)

    def finalize(self):
        par_content = '\n'.join(self.cur_paragraph_lines)
        par = self.state.doc.add_paragraph(par_content)
        par.style = self.state.doc.styles[self.style_name]

class LoadStyleHandler:
    Name = "load-style"

    def __init__(self, state: 'GdocxState', macro_args: list[str]):
        if len(macro_args) == 0:
            raise Exception("ParseStyleDirective macro must contain path to file")
        GdocxStyle.use_styles_from_file(macro_args[0], state.doc)

    def process_line(self, line: str, info: GdocxParsing.LineInfo):
        raise Exception("You must not place content inside ParseStyleDirective")

    def finalize(self):
        pass

class OrderedListHandler:
    Name = "ordered-list"

    def __init__(self, state: 'GdocxState', macro_args: list[str]):
        self.free_item_number = 1

    def process_line(self, line: str, info: GdocxParsing.LineInfo):
        raise Exception(f"You must provide contents of {self.Name} within macro {OrderedListItemHandler.Name}")

    def finalize(self):
        pass


class OrderedListItemHandler:
    Infix = ". "
    Name = "ordered-list-item"

    def __init__(self, state: 'GdocxState', macro_args: list[str]):
        if state.handler.Name != OrderedListHandler.Name:
            raise Exception(f"{self.Name} macro must be inside {OrderedListHandler.Name} macro")

        self.state = state
        self.cur_paragraph_lines = []
        self.is_first_line_processed = False

        self.item_number = state.handler.free_item_number
        state.handler.free_item_number += 1

    def process_line(self, line: str, info: GdocxParsing.LineInfo):
        if not self.is_first_line_processed:
            line_stripped = str(self.item_number) + self.Infix + info.line_stripped
            self.is_first_line_processed = True
        self.cur_paragraph_lines.append(line_stripped)

    def finalize(self):
        par_content = '\n'.join(self.cur_paragraph_lines)
        par = self.state.doc.add_paragraph(par_content)
        par.style = GdocxStyle.Style.UNORDERED_LIST

