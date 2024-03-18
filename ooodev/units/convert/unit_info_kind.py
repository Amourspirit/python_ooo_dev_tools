from __future__ import annotations
from enum import Enum


class UnitInfoKind(Enum):
    """Information units"""

    BIT_YOBI = "Yibit"
    """yobi Bit 2/80"""
    BIT_ZEBI = "Zibit"
    """zebi Bit 2/70"""
    BIT_EXBI = "Eibit"
    """exbi Bit 2/60"""
    BIT_PEBI = "Pibit"
    """pebi Bit 2/50"""
    BIT_TEBI = "Tibit"
    """tebi Bit 2/40"""
    BIT_GIBI = "Gibit"
    """gibi Bit 2/30"""
    BIT_MEBI = "Mibit"
    """mebi Bit 2/20"""
    BIT_KIBI = "kibit"
    """kibi Bit 2/10"""
    BIT = "bit"
    """Bit 2/0"""
    BYTE_YOBI = "Yibyte"
    """yobi Byte 2/80"""
    BYTE_ZEBI = "Zibyte"
    """zebi Byte 2/70"""
    BYTE_EXBI = "Eibyte"
    """exbi Byte 2/60"""
    BYTE_PEBI = "Pibyte"
    """pebi Byte 2/50"""
    BYTE_TEBI = "Tibyte"
    """tebi Byte 2/40"""
    BYTE_GIBI = "Gibyte"
    """gibi Byte 2/30"""
    BYTE_MEBI = "Mibyte"
    """mebi Byte 2/20"""
    BYTE_KIBI = "kibyte"
    """kibi Byte 2/10"""
    BYTE = "byte"
    """Byte 2/0"""

    def __str__(self) -> str:
        return self.value
