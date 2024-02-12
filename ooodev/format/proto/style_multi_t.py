from __future__ import annotations
from typing import TYPE_CHECKING
import uno

from ooodev.mock.mock_g import DOCS_BUILDING
from .style_t import StyleT

if TYPE_CHECKING or DOCS_BUILDING:
    try:
        from typing import Protocol
    except ImportError:
        from typing_extensions import Protocol
else:
    Protocol = object


class StyleMultiT(StyleT, Protocol):
    """Style Multi Base Protocol"""

    pass
