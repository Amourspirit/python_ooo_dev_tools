from enum import Flag


class StringKind(Flag):
    """Formatter String Kind Flags"""

    NONE = 0
    """No formatting"""
    LEFT_STRIP = 1 << 1
    """Strip Left"""


__all__ = ["StringKind"]
