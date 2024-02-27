from __future__ import annotations
from typing import Any, TYPE_CHECKING
from ooodev.mock import mock_g
from ooodev.adapter.chart2.data.data_provider_comp import DataProviderComp
from ooodev.loader import lo as mLo
from ooodev.office import chart2 as mChart2
from ooodev.exceptions import ex as mEx
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.calc.chart2.partial.chart_doc_prop_partial import ChartDocPropPartial
from ooodev.utils.context.lo_context import LoContext

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.calc.chart2.chart_doc import ChartDoc


class DataProvider(LoInstPropsPartial, DataProviderComp, ChartDocPropPartial):
    """
    Class for managing Chart2 Data Data Source.
    """

    def __init__(self, owner: ChartDoc, component: Any, lo_inst: LoInst | None = None) -> None:
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
        DataProviderComp.__init__(self, component)  # type: ignore
        ChartDocPropPartial.__init__(self, chart_doc=owner)

    def add_cat_labels(self, data_label: str, data_range: str) -> None:
        """
        Add Category Labels.

        |lo_unsafe|

        Args:
            chart_doc (XChartDocument): Chart Document.
            data_label (str): Data label.
            data_range (str): Data range.

        Raises:
            ChartError: If error occurs.

        Returns:
            None:
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.utils.kind.chart2_data_role_kind import DataRoleKind
        from ooodev.utils.kind.data_point_label_type_kind import DataPointLabelTypeKind

        try:
            dp = self.component
            with LoContext(self.lo_inst):
                dl_seq = mChart2.Chart2.create_ld_seq(
                    dp=dp, role=DataRoleKind.CATEGORIES, data_label=data_label, data_range=data_range
                )
            axis = self.chart_doc.axis_x
            sd = axis.component.getScaleData()
            sd.Categories = dl_seq
            axis.component.setScaleData(sd)

            # label the data points with these category values
            ds_arr = self.chart_doc.get_data_series()
            for ds in ds_arr:
                ds.set_data_point_labels(label_type=DataPointLabelTypeKind.CATEGORY)
        except mEx.ChartError:
            raise
        except Exception as e:
            raise mEx.ChartError("Error adding category labels") from e


if mock_g.FULL_IMPORT:
    from ooodev.utils.kind.chart2_data_role_kind import DataRoleKind
    from ooodev.utils.kind.data_point_label_type_kind import DataPointLabelTypeKind
