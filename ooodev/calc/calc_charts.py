from __future__ import annotations
from typing import TYPE_CHECKING
import uno

from ooodev.mock import mock_g
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
from ooodev.utils.kind.chart2_types import ChartTypes as ChartTypes
from ooodev.utils.partial.the_dictionary_partial import TheDictionaryPartial
from ooodev.office import chart2 as mCharts
from ooodev.utils.color import CommonColor
from ooodev.calc.chart2.table_chart import TableChart
from ooodev.calc.partial.calc_doc_prop_partial import CalcDocPropPartial
from ooodev.calc.partial.calc_sheet_prop_partial import CalcSheetPropPartial


if TYPE_CHECKING:
    from com.sun.star.table import XTableCharts
    from ooodev.utils.color import Color
    from ooodev.calc.calc_sheet import CalcSheet
    from ooodev.utils.kind.chart2_types import ChartTemplateBase


class CalcCharts(
    LoInstPropsPartial,
    TableChartsComp,
    QiPartial,
    ServicePartial,
    TheDictionaryPartial,
    CalcSheetPropPartial,
    CalcDocPropPartial,
):
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
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        TableChartsComp.__init__(self, component=charts)  # type: ignore
        QiPartial.__init__(self, component=charts, lo_inst=self.lo_inst)
        ServicePartial.__init__(self, component=charts, lo_inst=self.lo_inst)
        TheDictionaryPartial.__init__(self)
        CalcSheetPropPartial.__init__(self, obj=owner)
        CalcDocPropPartial.__init__(self, obj=owner.calc_doc)

    # region Iterable and Index Access
    def __next__(self) -> TableChart:
        """
        Gets the next chart.

        Returns:
            TableChart: The next chart.
        """
        return TableChart(owner=self.calc_sheet, component=super().__next__(), lo_inst=self.lo_inst)

    def __getitem__(self, index: str | int) -> TableChart:
        """
        Gets the chart at the specified index or name.

        This is short hand for ``get_by_index()`` or ``get_by_name()``.

        Args:
            key (key, str, int): The index or name of the chart. When getting by index can be a negative value to get from the end.

        Returns:
            TableChart: The chart with the specified index or name.

        See Also:
            - :py:meth:`~ooodev.calc.CalcCharts.get_by_index`
            - :py:meth:`~ooodev.calc.CalcCharts.get_by_name`
        """
        if isinstance(index, int):
            return self.get_by_index(index)
        return self.get_by_name(index)

    def __len__(self) -> int:
        """
        Gets the number of charts in the document.

        Returns:
            int: Number of charts in the document.
        """
        return self.component.getCount()

    def __delitem__(self, _item: int | str | TableChart) -> None:
        """
        Removes a chart from the document.

        Args:
            _item (int | str, TableChart): Index, name, or TableChart to remove.

        Raises:
            TypeError: If the item is not a supported type.
        """
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
        rng_obj: RangeObj | None = None,
        cell_name: str = "",
        width: int = 16,
        height: int = 9,
        diagram_name: ChartTemplateBase | str = "Column",
        color_bg: Color | None = CommonColor.PALE_BLUE,
        color_wall: Color | None = CommonColor.LIGHT_BLUE,
        **kwargs,
    ) -> TableChart:
        """
        Insert a new chart.

        Args:
            rng_obj (RangeObj, optional): Cell range object. Defaults to current selected cells.
            cell_name (str, optional): Cell name such as ``A1``.
            width (int, optional): Width. Default ``16``.
            height (int, optional): Height. Default ``9``.
            diagram_name (ChartTemplateBase | str): Diagram Name. Defaults to ``Column``.
            color_bg (:py:data:`~.utils.color.Color`, optional): Color Background. Defaults to ``CommonColor.PALE_BLUE``.
                If set to ``None`` then no color is applied.
            color_wall (:py:data:`~.utils.color.Color`, optional): Color Wall. Defaults to ``CommonColor.LIGHT_BLUE``.
                If set to ``None`` then no color is applied.

        Keyword Arguments:
            chart_name (str, optional): Chart name
            is_row (bool, optional): Determines if the data is row data or column data.
            first_cell_as_label (bool, optional): Set is first row is to be used as a label.
            set_data_point_labels (bool, optional): Determines if the data point labels are set.

        Raises:
            ChartError: If error occurs

        Returns:
            TableChart: Chart Document that was created and inserted into the sheet.

        Note:
            **Keyword Arguments** are to mostly be ignored.
            If finer control over chart creation is needed then **Keyword Arguments** can be used.

        Note:
            See **Open Office Wiki** - `The Structure of Charts <https://wiki.openoffice.org/wiki/Documentation/BASIC_Guide/Structure_of_Charts>`__ for more information.

        See Also:
            - :py:meth:`ooodev.office.chart2.Chart2.insert_chart`
            - :ref:`ooodev.utils.kind.chart2_types`
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.office.chart2 import Chart2

        # from ..utils.kind.chart2_types import ChartTemplateBase, ChartTypeNameBase, ChartTypes as ChartTypes
        if rng_obj is None:
            cr = None
        else:
            cr = rng_obj.get_cell_range_address()
        with LoContext(self.lo_inst):
            _ = Chart2.insert_chart(
                sheet=self.calc_sheet.component,
                cells_range=cr,
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


if mock_g.FULL_IMPORT:
    from ooodev.office.chart2 import Chart2
