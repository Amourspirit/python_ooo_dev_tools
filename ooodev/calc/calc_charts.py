from __future__ import annotations
from typing import TYPE_CHECKING
import uno

from ooodev.utils.context.lo_context import LoContext
from ooodev.exceptions import ex as mEx
from ooodev.adapter.table.table_charts_comp import TableChartsComp
from ooodev.utils import gen_util as mGenUtil
from ooodev.loader import lo as mLo
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.utils.data_type.range_obj import RangeObj
from ooodev.utils.kind.chart2_types import ChartTemplateBase, ChartTypes as ChartTypes
from ooodev.office import chart2 as mCharts
from ooodev.utils.color import CommonColor
from .chart2.table_chart import TableChart


if TYPE_CHECKING:
    from com.sun.star.table import XTableCharts
    from ooodev.utils.color import Color
    from .calc_sheet import CalcSheet


class CalcCharts(LoInstPropsPartial, TableChartsComp, QiPartial, ServicePartial):
    """
    Class for managing Calc Charts.

    .. versionadded:: 0.26.1.
    """

    def __init__(self, owner: CalcSheet, charts: XTableCharts, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (CalcDoc): Owner Document
            sheet (XSpreadsheet): Sheet instance.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self._owner = owner
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        TableChartsComp.__init__(self, component=charts)  # type: ignore
        QiPartial.__init__(self, component=charts, lo_inst=self.lo_inst)
        ServicePartial.__init__(self, component=charts, lo_inst=self.lo_inst)

    # region Iterable and Index Access
    def __next__(self) -> TableChart:
        return TableChart(owner=self._owner, component=super().__next__(), lo_inst=self.lo_inst)

    def __getitem__(self, index: str | int) -> TableChart:
        if isinstance(index, int):
            return self.get_by_index(index)
        return self.get_by_name(index)

    def __len__(self) -> int:
        return self.component.getCount()

    def __delitem__(self, _item: int | str | TableChart) -> None:
        # using remove_sheet here instead of remove_by_name. This will force Calc module event to be fired.
        if isinstance(_item, str):
            self.remove_by_name(_item)
        elif isinstance(_item, TableChart):
            self.remove_by_name(_item.name)
        elif isinstance(_item, int):
            tc = self.get_by_index(idx=_item)
            self.remove_by_name(tc.name)
        else:
            raise TypeError(f"Invalid type for __delitem__: {type(_item)}")

    def _get_index(self, idx: int, allow_greater: bool = False) -> int:
        """
        Gets the index.

        Args:
            idx (int): Index of sheet. Can be a negative value to index from the end of the list.
            allow_greater (bool, optional): If True and index is greater then the number of
                sheets then the index becomes the next index if sheet were appended. Defaults to False.

        Returns:
            int: Index value.
        """
        count = len(self)
        return mGenUtil.Util.get_index(idx, count, allow_greater)

    # endregion Iterable and Index Access

    # region XIndexAccess overrides

    def get_by_index(self, idx: int) -> TableChart:
        """
        Gets the element at the specified index.

        Args:
            idx (int): The Zero-based index of the element. Idx can be a negative value to index from the end of the list.
                For example, -1 will return the last element.

        Returns:
            CalcSheet: The element at the specified index.
        """
        index = self._get_index(idx)
        result = super().get_by_index(index)
        return TableChart(owner=self.calc_sheet, component=result, lo_inst=self.lo_inst)

    # endregion XIndexAccess overrides

    # region XNameAccess overrides

    def get_by_name(self, name: str) -> TableChart:
        """
        Gets the element with the specified name.

        Args:
            name (str): The name of the element.

        Raises:
            MissingNameError: If sheet is not found.

        Returns:
            CalcSheet: The element with the specified name.
        """
        if not self.has_by_name(name):
            raise mEx.MissingNameError(f"Unable to find sheet with name '{name}'")
        result = super().get_by_name(name)
        return TableChart(owner=self.calc_sheet, component=result, lo_inst=self.lo_inst)

    # endregion XNameAccess overrides

    def insert_chart(
        self,
        *,
        rng_obj: RangeObj,
        cell_name: str = "",
        width: int = 16,
        height: int = 9,
        diagram_name: ChartTemplateBase | str = "Column",
        color_bg: Color = CommonColor.PALE_BLUE,
        color_wall: Color = CommonColor.LIGHT_BLUE,
        **kwargs,
    ) -> TableChart:
        from ooodev.office.chart2 import Chart2

        # from ..utils.kind.chart2_types import ChartTemplateBase, ChartTypeNameBase, ChartTypes as ChartTypes
        with LoContext(self.lo_inst):
            _ = Chart2.insert_chart(
                sheet=self.calc_sheet.component,
                cells_range=rng_obj.get_cell_range_address(),
                cell_name=cell_name,
                width=width,
                height=height,
                diagram_name=diagram_name,
                color_bg=color_bg,
                color_wall=color_wall,
                **kwargs,
            )
        return self.get_by_index(-1)

    def remove_chart(self, chart_name: str) -> bool:
        """
        Removes a chart from Spreadsheet.

        Args:
            sheet (XSpreadsheet): Spreadsheet
            chart_name (str): Chart Name

        Returns:
            bool: ``True`` if chart was removed; Otherwise, ``False``
        """
        return mCharts.Chart2.remove_chart(sheet=self.calc_sheet.component, chart_name=chart_name)

    def clear(self) -> None:
        """
        Removes all charts from the sheet.
        """
        names = set()
        for chart in self:
            names.add(chart.name)
        for name in names:
            self.remove_chart(chart_name=name)

    # region Properties
    @property
    def calc_sheet(self) -> CalcSheet:
        """Sheet that owns this cell."""
        return self._owner

    # endregion Properties
