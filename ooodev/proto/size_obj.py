from __future__ import annotations
from typing import TYPE_CHECKING

from ..format.inner.kind.format_kind import FormatKind as FormatKind

if TYPE_CHECKING:
    try:
        from typing import Protocol
    except ImportError:
        from typing_extensions import Protocol
else:
    Protocol = object


class SizeObj(Protocol):
    """Protocol Class for size"""

    Width: int = ...
    """Size Width."""
    Height: int = ...
    """Size Height"""
