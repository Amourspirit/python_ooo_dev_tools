from enum import IntFlag


class FormatKind(IntFlag):
    """Format Kind"""

    UNKNOWN = 0
    PARA = 1 << 1
    CHAR = 1 << 2
    STRUCT = 1 << 3
    PARA_COMPLEX = 1 << 4
    CELL = 1 << 5
    STATIC = 1 << 6  # no backups if properties in write, etc.
    TXT_CONTENT = 1 << 7
    STYLE = 1 << 8
    PAGE = 1 << 9
    FILL = 1 << 10
    DOC = 1 << 11
    FRAME = 1 << 12
    SHAPE = 1 << 13
    IMAGE = 1 << 14
    TABLE = 1 << 15
    HEADER = 1 << 16
    FOOTER = 1 << 17
    """Document level Styles"""
