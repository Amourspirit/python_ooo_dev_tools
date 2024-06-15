from __future__ import annotations
from typing import cast, overload, TYPE_CHECKING
import uno


from ooodev.adapter.table.table_column_comp import TableColumnComp
from ooodev.format.inner.style_partial import StylePartial
from ooodev.office import calc as mCalc
from ooodev.units.unit_mm100 import UnitMM100
from ooodev.utils import info as mInfo
from ooodev.loader import lo as mLo
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.utils.partial.the_dictionary_partial import TheDictionaryPartial
from ooodev.calc.partial.calc_sheet_prop_partial import CalcSheetPropPartial
from ooodev.calc.partial.calc_doc_prop_partial import CalcDocPropPartial

if TYPE_CHECKING:
    from com.sun.star.table import CellAddress
    from com.sun.star.table import TableColumn  # service
    from com.sun.star.table import XCellRange
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.units.unit_obj import UnitT
    from ooodev.utils.data_type.cell_obj import CellObj
    from ooodev.utils.data_type.cell_values import CellValues
    from ooodev.utils.data_type.range_obj import RangeObj
    from ooodev.calc.calc_sheet import CalcSheet


class CalcTableCol(
    LoInstPropsPartial,
    TableColumnComp,
    QiPartial,
    ServicePartial,
    TheDictionaryPartial,
    StylePartial,
    CalcSheetPropPartial,
    CalcDocPropPartial,
):
    """Represents a calc table column."""

    def __init__(self, owner: CalcSheet, col_obj: TableColumn | int, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (CalcSheet): Sheet that owns this cell range.
            col_obj (Any): Range object.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        if mInfo.Info.is_instance(col_obj, int):
            comp = mCalc.Calc.get_col_range(sheet=self.calc_sheet.component, idx=col_obj)
            self._range_obj = mCalc.Calc.get_range_obj(cell_range=comp)
        else:
            self._range_obj = mCalc.Calc.get_range_obj(cell_range=cast("XCellRange", col_obj))
            comp = col_obj
        TableColumnComp.__init__(self, comp)  # type: ignore
        QiPartial.__init__(self, component=comp, lo_inst=self.lo_inst)  # type: ignore
        ServicePartial.__init__(self, component=comp, lo_inst=self.lo_inst)
        TheDictionaryPartial.__init__(self)
        StylePartial.__init__(self, component=comp)
        CalcSheetPropPartial.__init__(self, obj=owner)
        CalcDocPropPartial.__init__(self, obj=owner.calc_doc)
        # self.__doc = doc

    # region contains()

    @overload
    def contains(self, cell_obj: CellObj) -> bool:
        """
        Gets if current instance contains a cell value.

        Args:
            cell_obj (CellObj): Cell object

        Returns:
            bool: ``True`` if instance contains cell; Otherwise, ``False``.
        """
        ...

    @overload
    def contains(self, cell_addr: CellAddress) -> bool:
        """
        Gets if current instance contains a cell value.

        Args:
            cell_addr (CellAddress): Cell address

        Returns:
            bool: ``True`` if instance contains cell; Otherwise, ``False``.
        """
        ...

    @overload
    def contains(self, cell_vals: CellValues) -> bool:
        """
        Gets if current instance contains a cell value.

        Args:
            cell_vals (CellValues): Cell Values

        Returns:
            bool: ``True`` if instance contains cell; Otherwise, ``False``.
        """
        ...

    @overload
    def contains(self, cell_name: str) -> bool:
        """
        Gets if current instance contains a cell value.

        Args:
            cell_name (str): Cell name

        Returns:
            bool: ``True`` if instance contains cell; Otherwise, ``False``.
        """
        ...

    def contains(self, *args, **kwargs) -> bool:
        """
        Gets if current instance contains a cell value.

        Args:
            cell_obj (CellObj): Cell object
            cell_addr (CellAddress): Cell address
            cell_vals (CellValues): Cell Values
            cell_name (str): Cell name

        Returns:
            bool: ``True`` if instance contains cell; Otherwise, ``False``.

        Note:
            If cell input contains sheet info the it is use in comparison.
            Otherwise sheet is ignored.
        """
        return self.range_obj.contains(*args, **kwargs)

    # endregion contains()

    # region Properties

    @property
    def range_obj(self) -> RangeObj:
        """Range object."""
        return self._range_obj

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
