from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING, Union

import uno
import unohelper
from com.sun.star.sheet import XRangeSelectionChangeListener
from ooo.dyn.table.cell_content_type import CellContentType

from ooodev.io.log.named_logger import NamedLogger

if TYPE_CHECKING:
    from com.sun.star.sheet import RangeSelectionEvent
    from com.sun.star.table import XCell
    from com.sun.star.lang import EventObject
    from com.sun.star.sheet import SheetCell

    from com.sun.star.table import Cell
    from com.sun.star.table import CellProperties
    from com.sun.star.style import CharacterProperties
    from com.sun.star.style import ParagraphProperties
    from com.sun.star.sheet import SheetCellRange
    from com.sun.star.table import CellRange

    ScCellObj = Union[
        SheetCell, Cell, CellProperties, CharacterProperties, ParagraphProperties, SheetCellRange, CellRange
    ]
    ScCellRangeObj = Union[SheetCellRange, CellRange, CellProperties, CharacterProperties, ParagraphProperties]
else:
    ScCellObj = Any
    ScCellRangeObj = Any

# see: https://wiki.documentfoundation.org/Documentation/DevGuide/Spreadsheet_Documents#Range_Selection


class RangeSelectionChangeListener(unohelper.Base, XRangeSelectionChangeListener):

    def __init__(self) -> None:
        super().__init__()
        self._logger = NamedLogger("RangeSelectionChangeListener")
        self._prev_addr = None
        self._current_val = None

    def _convert_to_float(self, val: Any) -> float:
        """
        Converts value to float.

        |lo_safe|

        Args:
            val (Any): Value to convert

        Returns:
            float: value converted to float. 0.0 is returned if conversion fails.
        """
        if val is None:
            self._logger.debug("_convert_to_float(): Value is null; using 0")
            return 0.0
        try:
            return float(val)
        except ValueError:
            self._logger.debug(f"_convert_to_float(): Could not convert {val} to double; using 0")
            return 0.0

    def _get_val_by_cell(self, cell: XCell) -> Any:
        """LO Safe Method."""
        t = cell.getType()
        if t == CellContentType.EMPTY:
            return None
        if t == CellContentType.VALUE:
            return self._convert_to_float(cell.getValue())
        if t in (CellContentType.TEXT, CellContentType.FORMULA):
            return cell.getFormula()
        self._logger.debug("_get_val_by_cell(): Unknown cell type; returning None")
        return None

    def _is_prev_cell_val(self) -> bool:
        if self._prev_addr is None:
            return False
        if self._current_val is None:
            return False
        if hasattr(self._prev_addr, "typeName"):
            return getattr(self._prev_addr, "typeName") == "com.sun.star.table.CellAddress"
        return False

    def descriptorChanged(self, event: RangeSelectionEvent) -> None:
        """
        is called when the selected range is changed while range selection is active.
        """
        print(event.Source)

        # print(event.Source)

    def disposing(self, event: EventObject) -> None:
        """
        Gets called when the broadcaster is about to be disposed.

        All listeners and all other objects, which reference the broadcaster
        should release the reference to the source. No method should be invoked
        anymore on this object ( including XComponent.removeEventListener() ).

        This method is called for every listener registration of derived listener
        interfaced, not only for registrations at XComponent.
        """
        self._logger.debug("Disposing")
