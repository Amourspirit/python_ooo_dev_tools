from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING
import uno


from ooodev.adapter.sheet.spreadsheet_view_comp import SpreadsheetViewComp
from ooodev.adapter.sheet.spreadsheet_view_settings_comp import SpreadsheetViewSettingsComp
from ooodev.calc import calc_cell as mCalcCell
from ooodev.calc import calc_cell_cursor as mCalcCellCursor
from ooodev.calc import calc_cell_range as mCalcCellRange
from ooodev.calc.partial.calc_doc_prop_partial import CalcDocPropPartial
from ooodev.format.inner.style_partial import StylePartial
from ooodev.loader import lo as mLo
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils import info as mInfo
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.service_partial import ServicePartial

if TYPE_CHECKING:
    from com.sun.star.sheet import XSpreadsheetView

    # from com.sun.star.sheet import SpreadsheetView  # service
    # from com.sun.star.sheet import SpreadsheetViewSettings  # service
    from ooodev.calc.calc_doc import CalcDoc


class CalcSheetView(
    LoInstPropsPartial,
    SpreadsheetViewComp,
    SpreadsheetViewSettingsComp,
    QiPartial,
    PropPartial,
    StylePartial,
    ServicePartial,
    CalcDocPropPartial,
):
    def __init__(self, owner: CalcDoc, view: XSpreadsheetView, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        SpreadsheetViewComp.__init__(self, view)  # type: ignore
        SpreadsheetViewSettingsComp.__init__(self, view)  # type: ignore
        QiPartial.__init__(self, component=view, lo_inst=self.lo_inst)
        PropPartial.__init__(self, component=view, lo_inst=self.lo_inst)
        StylePartial.__init__(self, component=view)
        ServicePartial.__init__(self, component=view, lo_inst=self.lo_inst)
        CalcDocPropPartial.__init__(self, obj=owner.calc_doc)

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
