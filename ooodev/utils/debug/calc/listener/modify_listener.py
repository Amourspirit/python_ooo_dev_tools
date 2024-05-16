from __future__ import annotations
from typing import cast, TYPE_CHECKING, Any

import uno
import unohelper
from com.sun.star.util import XModifyListener
from com.sun.star.sheet import XSpreadsheetDocument

from ooodev.loader.lo import Lo
from ooodev.io.log.named_logger import NamedLogger
from ooo.dyn.table.cell_content_type import CellContentType

if TYPE_CHECKING:
    from com.sun.star.table import XCell
    from com.sun.star.lang import EventObject
    from com.sun.star.sheet import XCellRangeAddressable
    from com.sun.star.container import XIndexAccess
    from com.sun.star.sheet import SheetCell


class ModifyListener(unohelper.Base, XModifyListener):
    """
    Sheet Modify Listener.

    This listener is used to listen for changes in a sheet.
    This class is used for debugging purposes and can be attached to a sheet view.

    Example:
        .. code-block:: python

            with Lo.Loader(connector=Lo.ConnectPipe(), opt=Options(log_level=logging.DEBUG)):
                doc = CalcDoc.create_doc(visible=True)

                sheet = doc.sheets[0]

                sheet.set_col(
                    values=("Smith", 42, 58.9, -66.5, 43.4, 44.5, 45.3),
                    cell_name="A1",
                )

                ml = ModifyListener()
                doc.component.addModifyListener(ml)
    """

    def __init__(self) -> None:
        super().__init__()
        self._logger = NamedLogger("ModifyListener")

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

    def modified(self, event: EventObject) -> None:
        """
        is called when something changes in the object.

        Due to such an event, it may be necessary to update views or controllers.

        The source of the event may be the content of the object to which the listener
        is registered.
        """
        self._logger.debug("Modified")
        doc = Lo.qi(XSpreadsheetDocument, event.Source, True)

        sel = cast("XCellRangeAddressable", doc.getCurrentSelection())  # type: ignore
        addr = sel.getRangeAddress()  # Calc.get_selected_cell_addr(doc)
        self._logger.debug(
            f"Address: Sheet {addr.Sheet}, Start: {addr.StartRow},{addr.StartColumn}, End: {addr.EndRow},{addr.EndColumn}"
        )
        sheets = cast("XIndexAccess", doc.getSheets())
        sheet = sheets.getByIndex(addr.Sheet)
        cell = cast("SheetCell", sheet.getCellByPosition(addr.StartColumn, addr.StartRow))
        cell_val = self._get_val_by_cell(cell)
        formula = cell.getFormula()
        self._logger.debug("Formula: %s", formula)
        self._logger.debug("Cell Value %s", cell_val)

        # print(f"  {Calc.get_cell_str(addr=addr)} = {self._sheet.get_val(addr=addr)}")

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
