from enum import Flag


class TableItemKind(Flag):
    """Formatter Flags"""

    NONE = 0
    """No Options"""
    END_COL_RIGHT_STRIP = 1 << 0
    """Right Strip Last Column"""
    END_COL_LEFT_STRIP = 1 << 1
    """Left Strip Last Column"""
    START_COL_RIGHT_STRIP = 1 << 2
    """Right Strip Start Column"""
    START_COL_LEFT_STRIP = 1 << 3
    """Left Strip Start Column"""
    COL_RIGHT_STRIP = 1 << 4
    """Right Strip All Columns"""
    COL_LEFT_STRIP = 1 << 5
    """Left Strip All Columns"""
    START_COL_STRIP = START_COL_LEFT_STRIP | START_COL_RIGHT_STRIP
    """Strip Start Column"""
    END_COL_STRIP = END_COL_LEFT_STRIP | END_COL_RIGHT_STRIP
    """Strip End Column"""


__all__ = ["TableItemKind"]
