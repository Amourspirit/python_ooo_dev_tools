from enum import Flag


class FormatterTableRuleKind(Flag):
    IGNORE = 1 << 0
    ONLY = 1 << 1
    COL_OVER_ROW = 1 << 2
    CUSTOM_COL_OVER_ROW = 1 << 3
    RIGHT_STRIP_END_COL = 1 << 4
