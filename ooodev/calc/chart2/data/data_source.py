from __future__ import annotations
from typing import TYPE_CHECKING, Tuple
from ooodev.adapter.chart2.data.data_source_comp import DataSourceComp
from ooodev.loader import lo as mLo
from ooodev.office import chart2 as mChart2
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.calc.chart2.partial.chart_doc_prop_partial import ChartDocPropPartial


if TYPE_CHECKING:
    from com.sun.star.chart2.data import XDataSource
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.calc.chart2.chart_data_series import ChartDataSeries


class DataSource(LoInstPropsPartial, DataSourceComp, ChartDocPropPartial):
    """
    Class for managing Chart2 Data Data Source.
    """

    def __init__(self, owner: ChartDataSeries, component: XDataSource, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (ChartDiagram): Chart Diagram.
            component (XDataSource): UNO object that implements ``XDataSource`` interface.
            lo_inst (LoInst, optional): Lo Instance. Use when creating multiple documents. Defaults to None.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        DataSourceComp.__init__(self, component)  # type: ignore
        ChartDocPropPartial.__init__(self, chart_doc=owner.chart_doc)
        self.__owner = owner

    # region Properties

    def get_data(self, idx: int) -> Tuple[float, ...]:
        """
        Get Data

        Args:
            idx (int): Index.

        Raises:
            IndexError: If index is out of range.
            ChartError: If any other error occurs.

        Returns:
            Tuple[Tuple, ...]: float.
        """
        return mChart2.Chart2.get_chart_data(self.component, idx)

    @property
    def owner(self) -> ChartDataSeries:
        """Chart Diagram"""
        return self.__owner

    # endregion Properties
