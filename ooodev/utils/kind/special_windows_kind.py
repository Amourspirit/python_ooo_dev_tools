from __future__ import annotations
from enum import Enum


class SpecialWindowsKind(str, Enum):
    BASIC_IDE = "BASICIDE"
    WELCOME_SCREEN = "WELCOMESCREEN"

    def __str__(self) -> str:
        return self.value
