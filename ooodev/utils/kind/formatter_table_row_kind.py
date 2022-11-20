from enum import Flag


class FormatterTableRowKind(Flag):
    NONE = 0
    LEFT_STRIP = 1 << 1
