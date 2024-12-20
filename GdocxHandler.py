import GdocxParsing
import GdocxStyle
from docx.shared import Cm

# handler trait:
#   NAME: str
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
    NAME = "echo"

    def __init__(self, state: 'GdocxState', macro_args: list[str]):
        print('echo: ', *macro_args)
        pass

    def process_line(self, line: str, info: GdocxParsing.LineInfo):
        raise Exception("You must not place content inside echo")

    def finalize(self):
        pass

# Converts paragraphs to unordered list items
class UnorderedListItemHandler:
    NAME = "unordered-list-item"

    def __init__(self, state: 'GdocxState', macro_args: list[str]):
        self.state = state
        self.cur_paragraph_lines = []
        self.is_first_line_processed = False
        state.handler.finalize()

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
    NAME = "paragraph-styled"

    def __init__(self, state: 'GdocxState', macro_args: list[str]):
        if len(macro_args) == 0:
            raise Exception("Style macro must contain style name")
        self.state = state
        self.style_name = macro_args[0]
        self.cur_paragraph_lines = []
        state.handler.finalize()

    def process_line(self, line: str, info: GdocxParsing.LineInfo):
        self.cur_paragraph_lines.append(info.line_stripped)

    def finalize(self):
        par_content = '\n'.join(self.cur_paragraph_lines)
        par = self.state.doc.add_paragraph(par_content)
        par.style = self.state.doc.styles[self.style_name]

class LoadStyleHandler:
    NAME = "load-style"

    def __init__(self, state: 'GdocxState', macro_args: list[str]):
        if len(macro_args) == 0:
            raise Exception("ParseStyleDirective macro must contain path to file")
        GdocxStyle.use_styles_from_file(macro_args[0], state.doc)

    def process_line(self, line: str, info: GdocxParsing.LineInfo):
        raise Exception("You must not place content inside ParseStyleDirective")

    def finalize(self):
        pass

class OrderedListHandler:
    NAME = "ordered-list"

    def __init__(self, state: 'GdocxState', macro_args: list[str]):
        self.free_item_number = 1
        state.handler.finalize()

    def process_line(self, line: str, info: GdocxParsing.LineInfo):
        raise Exception(f"You must provide contents of {self.NAME} within macro {OrderedListItemHandler.NAME}")

    def finalize(self):
        pass


class OrderedListItemHandler:
    INFIX = ". "
    NAME = "ordered-list-item"

    def __init__(self, state: 'GdocxState', macro_args: list[str]):
        if state.handler.NAME != OrderedListHandler.NAME:
            raise Exception(f"{self.NAME} macro must be inside {OrderedListHandler.NAME} macro")

        self.state = state
        self.cur_paragraph_lines = []
        self.is_first_line_processed = False

        self.item_number = state.handler.free_item_number
        state.handler.free_item_number += 1

    def process_line(self, line: str, info: GdocxParsing.LineInfo):
        if not self.is_first_line_processed:
            line_stripped = str(self.item_number) + self.INFIX + info.line_stripped
            self.is_first_line_processed = True
        self.cur_paragraph_lines.append(line_stripped)

    def finalize(self):
        par_content = '\n'.join(self.cur_paragraph_lines)
        par = self.state.doc.add_paragraph(par_content)
        par.style = GdocxStyle.Style.ORDERED_LIST

class ImageHandler:
    NAME = "image"

    def __init__(self, state: 'GdocxState', macro_args: list[str]):
        if len(macro_args) == 0:
            raise Exception(f"Macro {self.NAME} called without arguments")
        self.state = state
        self.width = None
        self.height = None
        
        argslen = len(macro_args)
        if argslen > 1:
            widthstr = macro_args[1]
            if widthstr != "None":
                self.width = Cm(float(widthstr))

        if argslen > 2:
            heightstr = macro_args[2]
            if heightstr != "None":
                self.height = Cm(float(heightstr))

        self.path = macro_args[0]

    def process_line(self, line: str, info: GdocxParsing.LineInfo):
        raise Exception(f"You must not put contents into {self.NAME} macro")

    def finalize(self):
        par = self.state.doc.add_paragraph(None, style = GdocxStyle.Style.IMAGE_PARAGRAPH)
        run = par.add_run()
        run.add_picture(self.path, self.width, self.height)

class ImageCaptionHandler:
    # Because GOST wants us to minimize distance between image and its caption,
    # we insert the latter to the paragraph of the first. Thus, it's as if user
    # presses SHIFT+ENTER and then types caption by hand
    STICK_TO_PREV_PARAGRAPH = True
    NAME = "image-caption"
    ItemFreeNumber = 1

    def __init__(self, state: 'GdocxState', macro_args: list[str]):
        self.state = state
        self.paragraph_lines = []
        self.is_first_line_processed = False

        self.item_number = self.ItemFreeNumber
        self.ItemFreeNumber += 1

    def process_line(self, line: str, info: GdocxParsing.LineInfo):
        line_stripped = info.line_stripped
        if not self.is_first_line_processed:
            self.is_first_line_processed = True
            line_stripped = GdocxStyle.Style.IMAGE_CAPTION_PREFIX + str(self.item_number) + GdocxStyle.Style.IMAGE_CAPTION_INFIX + info.line_stripped

        self.paragraph_lines.append(line_stripped)

    def finalize(self):
        content = ' '.join(self.paragraph_lines)
        if not self.STICK_TO_PREV_PARAGRAPH:
            self.state.doc.add_paragraph(content, style = GdocxStyle.Style.IMAGE_CAPTION)
        else:
            content = '\n' + content
            self.state.doc.paragraphs[-1].add_run(content, style = GdocxStyle.Style.IMAGE_CAPTION)


IsHeadingWarningSet = False

class Heading1Handler:
    NAME = "heading-1"

    def __init__(self, state: 'GdocxState', macro_args: list[str]):
        global IsHeadingWarningSet
        if not IsHeadingWarningSet:
            IsHeadingWarningSet = True
            GdocxCommon.Warnings.append(GdocxCommon.GdocxWarning(state.lineno, "Can't construct TOC automatically. Add TOC manually"))

        self.state = state
        self.cur_paragraph_lines = []
        state.handler.finalize()

    def process_line(self, line: str, info: GdocxParsing.LineInfo):
        self.cur_paragraph_lines.append(info.line_stripped)

    def finalize(self):
        par_content = '\n'.join(self.cur_paragraph_lines)
        self.state.doc.add_paragraph(par_content, style = GdocxStyle.Style.HEADING_1)
        
class Heading2Handler:
    NAME = "heading-2"

    def __init__(self, state: 'GdocxState', macro_args: list[str]):
        global IsHeadingWarningSet
        if not IsHeadingWarningSet:
            IsHeadingWarningSet = True
            GdocxCommon.Warnings.append(GdocxCommon.GdocxWarning(state.lineno, "Can't construct TOC automatically. Add TOC manually"))

        self.state = state
        self.cur_paragraph_lines = []
        state.handler.finalize()

    def process_line(self, line: str, info: GdocxParsing.LineInfo):
        self.cur_paragraph_lines.append(info.line_stripped)

    def finalize(self):
        par_content = '\n'.join(self.cur_paragraph_lines)
        self.state.doc.add_paragraph(par_content, style = GdocxStyle.Style.HEADING_2)
