from __future__ import annotations
from typing import Tuple, TYPE_CHECKING
from ..format.kind.format_kind import FormatKind as FormatKind

try:
    from typing import Protocol
except ImportError:
    from typing_extensions import Protocol


class SizeObj(Protocol):
    """Protolcol Class for size"""

    Width: int = ...
    """Size Width."""
    Height: int = ...
    """Size Height"""
