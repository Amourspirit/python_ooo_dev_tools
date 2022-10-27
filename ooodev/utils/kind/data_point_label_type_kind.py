from enum import IntEnum


class DataPointLabelTypeKind(IntEnum):
    """Data Point Label Type Kind"""

    NUMBER = 0
    PERCENT = 1
    CATEGORY = 2
    SYMBOL = 3
    NONE = 4
