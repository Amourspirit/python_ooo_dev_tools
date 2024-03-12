from __future__ import annotations

from typing import Any, TYPE_CHECKING

if TYPE_CHECKING:
    from ooodev.draw.draw_doc import DrawDoc
else:
    DrawDoc = Any


class DrawDocPropPartial:
    """A partial class for Draw Document."""

    def __init__(self, obj: DrawDoc) -> None:
        self.__draw_doc = obj

    @property
    def draw_doc(self) -> DrawDoc:
        """Write Document."""
        return self.__draw_doc
