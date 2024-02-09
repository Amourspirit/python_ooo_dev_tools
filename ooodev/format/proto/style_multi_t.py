from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

if TYPE_CHECKING:
    try:
        from typing import Protocol
    except ImportError:
        from typing_extensions import Protocol
    from .style_t import StyleT
else:
    Protocol = object
    StyleT = Any


class StyleMultiT(StyleT, Protocol):
    """Style Multi Base Protocol"""

    pass
