from enum import IntEnum


class ShapeCombKind(IntEnum):
    """Shape combine Kind"""

    MERGE = 0
    INTERSECT = 1
    SUBTRACT = 2
    COMBINE = 3
