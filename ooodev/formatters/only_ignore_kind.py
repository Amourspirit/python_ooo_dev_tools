from enum import IntEnum


class OnlyIgnoreKind(IntEnum):
    NONE = 0
    ONLY = 1
    IGNORE = 2


__all__ = ["OnlyIgnoreKind"]
