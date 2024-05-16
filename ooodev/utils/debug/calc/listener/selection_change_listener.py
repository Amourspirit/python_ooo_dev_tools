from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING, Union

import uno
import unohelper
from com.sun.star.view import XSelectionChangeListener
from ooo.dyn.table.cell_content_type import CellContentType

from ooodev.io.log.named_logger import NamedLogger

if TYPE_CHECKING:
    from com.sun.star.table import XCell
    from com.sun.star.lang import EventObject
    from com.sun.star.sheet import SheetCell
    from com.sun.star.sheet import SpreadsheetView
    from com.sun.star.sheet import SpreadsheetViewSettings

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


class SelectionChangeListener(unohelper.Base, XSelectionChangeListener):
    """
    Sheet Selection Change Listener.

    This listener is used to listen for selection changes in a sheet.
    This class is used for debugging purposes and can be attached to a sheet view.

    Note:
        If a value is changed while it is selected this class will not be notified.

    Example:
        .. code-block:: python

            with Lo.Loader(connector=Lo.ConnectPipe(), opt=Options(log_level=logging.DEBUG)):
                doc = CalcDoc.create_doc(visible=True)

                sheet = doc.sheets[0]

                sheet.set_col(
                    values=("Smith", 42, 58.9, -66.5, 43.4, 44.5, 45.3),
                    cell_name="A1",
                )
                sel_listener = SelectionChangeListener()
                view = doc.get_view()
                view.add_selection_change_listener(sel_listener)
    """

    def __init__(self) -> None:
        super().__init__()
        self._logger = NamedLogger("SelectionChangeListener")
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

    def selectionChanged(self, event: EventObject) -> None:
        """
        Is called when the selection changes.

        You can get the new selection via XSelectionSupplier from com.sun.star.lang.EventObject.Source.
        """
        # is fired four times for every click, and twice for shift arrow keys (?)
        # Once for arrow keys.
        # Will fire as new cells are selected with a mouse drag.
        # Does not have a way to get the previous selection.
        # Does not notify when the selection is complete in a drag.
        # Maybe would have to be combines with mouse events.
        src = cast("Union[SpreadsheetView, SpreadsheetViewSettings]", event.Source)
        sheet = src.getActiveSheet()  # type: ignore

        sel = src.getSelection()  # type: ignore
        if sel is None:
            return
        name = sel.getImplementationName()
        if name == "ScCellObj":
            sel_cell = cast(ScCellObj, sel)
            # should get the array formula su as {={1,2, 3}} when it exist.
            # sel_cell.getArrayFormula()
            addr = sel_cell.getCellAddress()  # type: ignore
            if self._prev_addr and addr == self._prev_addr:
                self._logger.debug("Same address: %i, %i", addr.Row, addr.Column)
                return
            if self._is_prev_cell_val():
                self._logger.debug("Prev address: %i, %i", self._prev_addr.Row, self._prev_addr.Column)  # type: ignore
                self._logger.debug(f"Prev Val: {self._current_val}")
            else:
                self._logger.debug("No Prev Value")
            self._current_val = self._get_val_by_cell(sel_cell)  # type: ignore
            self._logger.debug(f"Current Val: {self._current_val}")
            self._logger.debug("address: %i, %i", addr.Row, addr.Column)
            formula = sel_cell.getFormula()  # type: ignore
            if formula and formula.startswith("="):
                self._logger.debug("Formula: %s", formula)
            self._prev_addr = addr
            return
        if name == "ScCellRangeObj":
            sel_cells = cast(ScCellRangeObj, sel)
            rng_addr = sel_cells.getRangeAddress()  # type: ignore
            if self._prev_addr and rng_addr == self._prev_addr:
                self._logger.debug(
                    "Same Range: %i, %i, %i, %i",
                    rng_addr.StartRow,
                    rng_addr.StartColumn,
                    rng_addr.EndRow,
                    rng_addr.EndColumn,
                )
                return
            self._logger.debug(
                "Range :  %i, %i, %i, %i", rng_addr.StartRow, rng_addr.StartColumn, rng_addr.EndRow, rng_addr.EndColumn
            )
            formula = sel_cells.getArrayFormula()  # type: ignore
            if formula and formula.startswith("{="):
                self._logger.debug("Array Formula: %s", formula)
            self._prev_addr = rng_addr
        # print(sel)
        self._prev_addr = None

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
