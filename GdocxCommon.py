
class GdocxWarning:
    FMT = "WARNING: line %d: %s"
    def __init__(self, msg: str, lineno: int):
        self.string = self.FMT % (lineno, msg)

    def __str__():
        return self.string

Warnings: list[GdocxWarning] = []
