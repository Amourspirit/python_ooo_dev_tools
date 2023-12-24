from __future__ import annotations
from typing import TYPE_CHECKING
import uno


if TYPE_CHECKING:
    from .calc_sheet import CalcSheet
    from ooodev.utils.data_type.range_obj import RangeObj
    from ooodev.units import UnitT

from ooodev.adapter.table.table_column_comp import TableColumnComp
from ooodev.format.inner.style_partial import StylePartial
from ooodev.office import calc as mCalc
from ooodev.units import UnitMM100
from ooodev.utils import lo as mLo
from ooodev.utils.partial.qi_partial import QiPartial


if TYPE_CHECKING:
    from com.sun.star.table import TableColumn  # service


class CalcTableCol(TableColumnComp, QiPartial, StylePartial):
    """Represents a calc table column."""

    def __init__(self, owner: CalcSheet, col_obj: TableColumn | int) -> None:
        """
        Constructor

        Args:
            owner (CalcSheet): Sheet that owns this cell range.
            col_obj (Any): Range object.
        """
        self.__owner = owner
        if isinstance(col_obj, int):
            comp = mCalc.Calc.get_col_range(sheet=self.calc_sheet.component, idx=col_obj)
            self.__range_obj = mCalc.Calc.get_range_obj(cell_range=comp)
        else:
            self.__range_obj = mCalc.Calc.get_range_obj(cell_range=col_obj)
            comp = col_obj
        TableColumnComp.__init__(self, comp)  # type: ignore
        QiPartial.__init__(self, component=comp, lo_inst=mLo.Lo.current_lo)  # type: ignore
        StylePartial.__init__(self, component=comp)
        # self.__doc = doc

    # region Properties
    @property
    def calc_sheet(self) -> CalcSheet:
        """Sheet that owns this cell."""
        return self.__owner

    @property
    def range_obj(self) -> RangeObj:
        """Range object."""
        return self.__range_obj

    @property
    def width(self) -> UnitMM100:
        """
        Gets/Sets the width of the column in ``1/100 mm``.

        The width of the column can be set with int in ``1/100 mm`` or any ``UnitT`` object.

        Returns:
            UnitMM100: Width in ``1/100 mm``.
        """
        return UnitMM100(self.component.Width)

    @width.setter
    def width(self, value: UnitT | int) -> None:
        try:
            _width = value.get_value_mm100()  # type: ignore
        except AttributeError:
            _width = value

        self.component.Width = _width  # type: ignore

    @property
    def is_visible(self) -> bool:
        """Gets/Sets the visibility of the column."""
        return self.component.IsVisible

    @is_visible.setter
    def is_visible(self, value: bool) -> None:
        self.component.IsVisible = value

    @property
    def optimal_width(self) -> bool:
        """
        Gets/Sets the optimal width of the column.

        If ``True``, the column always keeps its optimal width.
        """
        return self.component.OptimalWidth

    @optimal_width.setter
    def optimal_width(self, value: bool) -> None:
        self.component.OptimalWidth = value

    @property
    def is_start_of_new_page(self) -> bool:
        """
        Gets the start of new page of the column.

        If ``True``, there is a manual horizontal page break attached to the column.
        """
        return self.component.IsStartOfNewPage

    # endregion Properties
