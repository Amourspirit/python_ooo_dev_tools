from __future__ import annotations
from typing import Any, TYPE_CHECKING

if TYPE_CHECKING:
    from ooodev.calc.calc_doc import CalcDoc
else:
    CalcDoc = Any


class CalcDocPropPartial:
    """A partial class for Calc Document."""

    def __init__(self, obj: CalcDoc) -> None:
        self.__calc_doc = obj

    @property
    def calc_doc(self) -> CalcDoc:
        """Calc Document."""
        return self.__calc_doc
