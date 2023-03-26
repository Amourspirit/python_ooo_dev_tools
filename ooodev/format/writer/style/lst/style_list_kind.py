from enum import Enum


class StyleListKind(Enum):
    """Style Look ups for paragraph List Styles"""

    NONE = ""
    """No List Style"""
    LIST_01 = "List 1"
    """Bullet •"""
    LIST_02 = "List 2"
    """Dash –"""
    LIST_03 = "List 3"
    """Bullet 🗹 (checkbox like)"""
    LIST_04 = "List 4"
    """Bullet ‣ (triangle like)"""
    LIST_05 = "List 5"
    """Bullet ꭗ"""
    NUM_123 = "Numbering 123"
    """Numbering ``123``"""
    NUM_abc = "Numbering abc"
    """Numbering ``abc`` (lower case)"""
    NUM_ABC = "Numbering ABC"
    """Numbering ``ABC`` (upper case)"""
    NUM_ivx = "Numbering ivx"
    """Numbering ``ivx`` (lower case)"""
    NUM_IVX = "Numbering IVX"
    """Numbering ``IVX`` (upper case)"""

    def __str__(self) -> str:
        return self.value
