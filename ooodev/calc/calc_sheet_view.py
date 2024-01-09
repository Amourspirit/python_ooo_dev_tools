from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING, Union
import uno


from ooodev.adapter.sheet.spreadsheet_view_comp import SpreadsheetViewComp
from ooodev.format.inner.style_partial import StylePartial
from ooodev.utils import lo as mLo
from ooodev.utils import info as mInfo
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.adapter.sheet.spreadsheet_view_settings_comp import SpreadsheetViewSettingsComp
from . import calc_cell_range as mCalcCellRange
from . import calc_cell as mCalcCell
from . import calc_cell_cursor as mCalcCellCursor

if TYPE_CHECKING:
    from com.sun.star.sheet import XSpreadsheetView

    # from com.sun.star.sheet import SpreadsheetView  # service
    # from com.sun.star.sheet import SpreadsheetViewSettings  # service
    from .calc_doc import CalcDoc


class CalcSheetView(SpreadsheetViewComp, SpreadsheetViewSettingsComp, QiPartial, PropPartial, StylePartial):
    def __init__(self, owner: CalcDoc, view: XSpreadsheetView) -> None:
        self.__owner = owner
        SpreadsheetViewComp.__init__(self, view)  # type: ignore
        SpreadsheetViewSettingsComp.__init__(self, view)  # type: ignore
        QiPartial.__init__(self, component=view, lo_inst=mLo.Lo.current_lo)
        PropPartial.__init__(self, component=view, lo_inst=mLo.Lo.current_lo)
        StylePartial.__init__(self, component=view)

    @overload
    def select(self, selection: mCalcCellRange.CalcCellRange) -> bool:
        """
        Selects the cells represented by the selection.

        Args:
            selection (CalcCellRange): Selection

        Returns:
            bool: True if selection was successful; Otherwise, False
        """
        ...

    @overload
    def select(self, selection: mCalcCell.CalcCell) -> bool:
        """
        Selects the cells represented by the selection.

        Args:
            selection (CalcCell): Selection

        Returns:
            bool: True if selection was successful; Otherwise, False
        """
        ...

    @overload
    def select(self, selection: mCalcCellCursor.CalcCellCursor) -> bool:
        """
        Selects the cells represented by the selection.

        Args:
            selection (CalcCellCursor): Selection

        Returns:
            bool: True if selection was successful; Otherwise, False
        """
        ...

    def select(self, selection: Any) -> bool:
        """
        Selects the object represented by xSelection if it is known and selectable in this object.

        Args:
            selection (Any): Selection

        Returns:
            bool: True if selection was successful; Otherwise, False
        """
        if mInfo.Info.is_instance(selection, mCalcCellRange.CalcCellRange):
            cursor = selection.create_cursor()
            return self.component.select(cursor.component)
        elif mInfo.Info.is_instance(selection, mCalcCell.CalcCell):
            cursor = selection.create_cursor()
            return self.component.select(cursor.component)
        elif mInfo.Info.is_instance(selection, mCalcCellCursor.CalcCellCursor):
            cell_rng = selection.get_calc_cell_range()
            cursor = cell_rng.create_cursor()
            return self.component.select(cursor.component)

        return self.component.select(selection)

    def get_selection(self) -> Any:
        """
        Returns the current selection.

        Returns:
            Any: Selection
        """
        return self.component.getSelection()

    # region Properties
    # if TYPE_CHECKING:

    #     @property
    #     def component(self) -> Union[SpreadsheetView, SpreadsheetViewSettings]:
    #         """Spreadsheet View Component"""
    #         return super().component  # type: ignore

    @property
    def calc_doc(self) -> CalcDoc:
        """
        Returns:
            CalcDoc: Calc doc
        """
        return self.__owner

    # endregion Properties
