from __future__ import annotations
from enum import Flag


class BorderKind(Flag):
    """Border Kind"""

    NONE = 0
    LEFT = 1 << 0
    TOP = 1 << 1
    RIGHT = 1 << 2
    BOTTOM = 1 << 3
    ALL = LEFT | TOP | RIGHT | BOTTOM
