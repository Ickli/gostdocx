# Can modify
MACRO_START = "{"
ESCAPE_CHAR = "\\"
MACRO_END = "}"
HEADER_START = "#"

# Don't modify
MACRO_START_ESCAPED = ESCAPE_CHAR + MACRO_START
ESCAPE_CHAR_ESCAPED = ESCAPE_CHAR + MACRO_END

INFO_TYPE_PLAIN_LINE = 0
INFO_TYPE_MACRO = 1
INFO_TYPE_HEADER = 2

class LineInfo:
    def __init__(self, line: str):
        line = line.lstrip(" ")
        self.line_stripped = line

        if is_macro(line):
            self.type = INFO_TYPE_MACRO
        elif is_header(line):
            self.type = INFO_TYPE_HEADER
        else:
            self.type = INFO_TYPE_PLAIN_LINE

def process_line(line: str) -> (str, LineInfo):
    line = line.rstrip('\n')
    return line, LineInfo(line)

def is_macro(line) -> bool:
    return line.startswith(MACRO_START) or line.startswith(MACRO_END)

def is_header(line) -> bool:
    return line.startswith(HEADER_START)

def get_header_string(line) -> str:
    if not is_header(line):
        raise Exception(f"get_header_string: line \"{line}\" is not header")
    return line[len(HEADER_START):].lstrip(' ')

MACRO_TYPE_START = 0
MACRO_TYPE_END = 1
MACRO_TYPE_ONE_LINE = 2

def get_macro_type(line) -> int:
    if line.startswith(MACRO_START):
        if line.endswith(MACRO_END):
            return MACRO_TYPE_ONE_LINE
        return MACRO_TYPE_START
    if line.startswith(MACRO_END):
        return MACRO_TYPE_END
    raise Exception(f'"{line}" is not macro')

def parse_macro_args(line) -> list[str]:
    if not is_macro(line):
        raise Exception("gdocx_parse_macro_args: line is not macro")

    line = line[len(MACRO_START):]
    if line.endswith(MACRO_END):
        line = line[:-len(MACRO_END)]

    args = []

    in_quotes = False
    skipping_spaces = True
    spaces = " \t\v"
    word_start = 0
    line_len = len(line)

    for i in range(line_len):
        char = line[i]
        
        if char == '"':
            skipping_spaces = False
            if not in_quotes:
                word_start = i + 1
            else:
                args.append(line[word_start:i])
                skipping_spaces = True
            in_quotes = not in_quotes
        elif char in spaces and not in_quotes:
            if skipping_spaces:
                continue
            args.append(line[word_start:i])
            skipping_spaces = True
        else:
            if skipping_spaces:
                skipping_spaces = False
                word_start = i
            if i == line_len - 1:
                args.append(line[word_start:line_len])
    return args
