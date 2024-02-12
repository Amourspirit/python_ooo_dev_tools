from __future__ import annotations
from typing import Any, TYPE_CHECKING

import uno


if TYPE_CHECKING:
    try:
        from typing import Protocol
    except ImportError:
        from typing_extensions import Protocol
    from .abstract_fill_color_t import AbstractFillColor
else:
    Protocol = object
    AbstractFillColor = Any


class FillColorT(AbstractFillColor, Protocol):
    """Font Effect Protocol"""

    @property
    def empty(self) -> FillColorT:  # type: ignore[misc]
        """Gets FillColor empty."""
        ...
