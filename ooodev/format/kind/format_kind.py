from enum import IntFlag


class FormatKind(IntFlag):
    """Format Kind"""

    UNKNOWN = 0
    PARA = 0x01
    CHAR = 0x02
    STRUCT = 0x03
    PARA_COMPLEX = 0x04
    CELL = 0x05
