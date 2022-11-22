from enum import Flag


class TableRuleKind(Flag):
    IGNORE = 1 << 0
    """Do no format indexes"""
    ONLY = 1 << 1
    """Only format indexes"""
    COL_OVER_ROW = 1 << 2
    """Column takes priority over row when there is conflict"""
    RIGHT_STRIP_END_COL = 1 << 3
    """Right strip End Column"""


__all__ = ["TableRuleKind"]
