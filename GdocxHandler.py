import GdocxParsing
import GdocxStyle

class EchoHandler:
    def __init__(self, state: 'GdocxState', macro_args: list[str]):
        print('echo: ', *macro_args)
        pass

    def process_line(self, line: str, info: GdocxParsing.LineInfo):
        raise Exception("You must not place content inside echo")

    def finalize(self):
        pass

# Converts paragraphs to unordered list items
class UnorderedListHandler:
    def __init__(self, state: 'GdocxState', macro_args: list[str]):
        self.state = state
        self.cur_paragraph_lines = []

    def process_line(self, line: str, info: GdocxParsing.LineInfo):
        self.cur_paragraph_lines.append(line)

    def finalize(self):
        par_content = '\n'.join(line for line in self.cur_paragraph_lines)
        par = self.state.doc.add_paragraph(par_content)
        par.style = GdocxStyle.Style.UNORDERED_LIST

# Applies style to macro contents, treating it as a paragraph
class ParStyleHandler:
    def __init__(self, state: 'GdocxState', macro_args: list[str]):
        if len(macro_args) == 0:
            raise Exception("Style macro must contain style name")
        self.state = state
        self.style_name = macro_args[1]

    def process_line(self, line: str, info: GdocxParsing.LineInfo):
        self.cur_paragraph_lines.append(line)

    def finalize(self):
        par_content = '\n'.join(line for line in self.cur_paragraph_lines)
        par = self.state.doc.add_paragraph(par_content)
        par.style = self.state.doc.styles[self.style_name]

class LoadStyleHandler:
    def __init__(self, state: 'GdocxState', macro_args: list[str]):
        if len(macro_args) == 0:
            raise Exception("ParseStyleDirective macro must contain path to file")
        file = open(macro_args[0], "r")
        json_string = file.read()
        GdocxStyle.parse_raw_styles(json_string, state.doc)
        print('loaded sth')

    def process_line(self, line: str, info: GdocxParsing.LineInfo):
        raise Exception("You must not place content inside ParseStyleDirective")

    def finalize(self):
        pass
