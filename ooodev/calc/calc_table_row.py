from __future__ import annotations
from typing import TYPE_CHECKING
import uno


if TYPE_CHECKING:
    from .calc_sheet import CalcSheet
    from ooodev.utils.data_type.range_obj import RangeObj
    from ooodev.units import UnitT

from ooodev.office import calc as mCalc
from ooodev.adapter.table.table_row_comp import TableRowComp
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils import lo as mLo
from ooodev.units import UnitMM100

if TYPE_CHECKING:
    from com.sun.star.table import TableColumn  # service


class CalcTableRow(TableRowComp, QiPartial):
    """Represents a calc table row."""

    def __init__(self, owner: CalcSheet, row_obj: TableColumn | int) -> None:
        """
        Constructor

        Args:
            owner (CalcSheet): Sheet that owns this cell range.
            TableColumn | int (Any): Range object.
        """
        self.__owner = owner
        if isinstance(row_obj, int):
            comp = mCalc.Calc.get_row_range(sheet=self.calc_sheet.component, idx=row_obj)
            self.__range_obj = mCalc.Calc.get_range_obj(cell_range=comp)
        else:
            self.__range_obj = mCalc.Calc.get_range_obj(cell_range=row_obj)
            comp = row_obj
        TableRowComp.__init__(self, comp)  # type: ignore
        QiPartial.__init__(self, component=comp, lo_inst=mLo.Lo.current_lo)  # type: ignore
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
    def height(self) -> UnitMM100:
        """
        Gets/Sets the height of the row in ``1/100 mm``.

        The height of the row can be set with int in ``1/100 mm`` or any ``UnitT`` object.

        Returns:
            UnitMM100: Width in ``1/100 mm``.
        """
        return UnitMM100(self.component.Height)

    @height.setter
    def height(self, value: UnitT | int) -> None:
        try:
            _height = value.get_value_mm100()  # type: ignore
        except AttributeError:
            _height = value

        self.component.Height = _height  # type: ignore

    @property
    def is_visible(self) -> bool:
        """Gets/Sets the visibility of the row."""
        return self.component.IsVisible

    @is_visible.setter
    def is_visible(self, value: bool) -> None:
        self.component.IsVisible = value

    @property
    def optimal_height(self) -> bool:
        """
        Gets/Sets the optimal height of the row.

        If ``True``, the row always keeps its optimal height.
        """
        return self.component.OptimalHeight

    @optimal_height.setter
    def optimal_height(self, value: bool) -> None:
        self.component.OptimalHeight = value

    @property
    def is_start_of_new_page(self) -> bool:
        """
        Gets the start of new page of the row.

        If ``True``, there is a manual vertical page break attached to the row..
        """
        return self.component.IsStartOfNewPage

    # endregion Properties
