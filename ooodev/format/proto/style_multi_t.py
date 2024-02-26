from __future__ import annotations
from typing import TYPE_CHECKING
import uno

from ooodev.mock.mock_g import DOCS_BUILDING
from ooodev.format.proto.style_t import StyleT

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Protocol
else:
    Protocol = object


class StyleMultiT(StyleT, Protocol):
    """Style Multi Base Protocol"""

    pass
