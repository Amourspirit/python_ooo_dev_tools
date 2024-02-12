from __future__ import annotations

from typing import Any, TYPE_CHECKING

if TYPE_CHECKING:
    from ..chart_doc import ChartDoc
else:
    ChartDoc = Any


class ChartDocPropPartial:
    """A partial class for Chart Document."""

    def __init__(self, chart_doc: ChartDoc) -> None:
        self.__chart_doc = chart_doc

    @property
    def chart_doc(self) -> ChartDoc:
        """Chart Document."""
        return self.__chart_doc
