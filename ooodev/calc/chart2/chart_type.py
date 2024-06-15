from __future__ import annotations

from typing import Any, TYPE_CHECKING, TypeVar, Generic, Tuple

from ooodev.mock import mock_g
from ooodev.adapter.chart2.chart_type_comp import ChartTypeComp
from ooodev.calc.chart2.partial.chart_doc_prop_partial import ChartDocPropPartial
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils.comp.prop import Prop
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.utils.partial.the_dictionary_partial import TheDictionaryPartial
from ooodev.calc.partial.calc_doc_prop_partial import CalcDocPropPartial
from ooodev.calc.partial.calc_sheet_prop_partial import CalcSheetPropPartial
from ooodev.utils import color as mColor
from ooodev.proto.component_proto import ComponentT

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.calc.chart2.chart_data_series import ChartDataSeries
    from ooodev.calc.chart2.chart_doc import ChartDoc
else:
    CoordinateGeneral = Any

_T = TypeVar("_T", bound="ComponentT")


class ChartType(
    Generic[_T],
    LoInstPropsPartial,
    ChartTypeComp,
    EventsPartial,
    ChartDocPropPartial,
    PropPartial,
    QiPartial,
    ServicePartial,
    TheDictionaryPartial,
    CalcDocPropPartial,
    CalcSheetPropPartial,
):
    """
    Class for managing Chart2 Chart Title Component.
    """

    def __init__(self, owner: _T, chart_doc: ChartDoc, component: Any, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Chart2 Title Component.
            lo_inst (LoInst, optional): Lo Instance. Use when creating multiple documents. Defaults to None.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        ChartTypeComp.__init__(self, component=component)
        EventsPartial.__init__(self)
        ChartDocPropPartial.__init__(self, chart_doc=chart_doc)
        PropPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        ServicePartial.__init__(self, component=component, lo_inst=self.lo_inst)
        TheDictionaryPartial.__init__(self)
        CalcDocPropPartial.__init__(self, obj=chart_doc.calc_doc)
        CalcSheetPropPartial.__init__(self, obj=chart_doc.calc_sheet)
        self._owner = owner
        self.get_data_series()

    # region DataSeriesContainerPartial overrides
    def get_data_series(self) -> Tuple[ChartDataSeries[ChartType[_T]], ...]:
        """
        retrieve all data series
        """
        # pylint: disable=import-outside-toplevel
        from .chart_data_series import ChartDataSeries

        d_series = super().get_data_series()
        if not d_series:
            return ()
        return tuple(ChartDataSeries(owner=self, chart_doc=self.chart_doc, component=ds) for ds in d_series)

    # endregion DataSeriesContainerPartial overrides

    def color_stock_bars(self, white_day_color: mColor.Color, black_day_color: mColor.Color) -> None:
        """
        Set color of stock bars for a ``CandleStickChartType`` chart.

        Args:
            white_day_color (~ooodev.utils.color.Color): Chart white day color
            black_day_color (~ooodev.utils.color.Color): Chart black day color

        Raises:
            NotSupportedError: If Chart is not of type ``CandleStickChartType``
            ChartError: If any other error occurs.

        Returns:
            None:

        See Also:
            :py:class:`~.color.CommonColor`
        """
        if self.chart_type != "com.sun.star.chart2.CandleStickChartType":
            raise mEx.NotSupportedError(f'Only candle stick charts supported. "{self.chart_type}" not supported.')
        try:
            white_day_ps = Prop(owner=self, component=self.get_property("WhiteDay"), lo_inst=self.lo_inst)
            black_day_ps = Prop(owner=self, component=self.get_property("BlackDay"), lo_inst=self.lo_inst)
            white_day_ps.set_property(FillColor=int(white_day_color))
            black_day_ps.set_property(FillColor=int(black_day_color))
        except Exception as e:
            raise mEx.ChartError("Error coloring stock bars") from e

    @property
    def owner(self) -> _T:
        """Gets Owner object for this Chart Type"""
        return self._owner

    @property
    def chart_type(self) -> str:
        """
        Gets chart type such as ``com.sun.star.chart2.StockBarChart``.
        """
        return self.get_chart_type()


if mock_g.FULL_IMPORT:
    from .chart_data_series import ChartDataSeries
