from __future__ import annotations
from typing import Any, TYPE_CHECKING

import uno
from ooodev.mock.mock_g import DOCS_BUILDING
from ooodev.format.proto.area.abstract_fill_color_t import AbstractFillColor

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Self
    from typing_extensions import Protocol
else:
    Protocol = object
    Self = Any


class FillColorT(AbstractFillColor, Protocol):
    """Font Effect Protocol"""

    @property
    def empty(self) -> Self:  # type: ignore[misc]
        """Gets FillColor empty."""
        ...
