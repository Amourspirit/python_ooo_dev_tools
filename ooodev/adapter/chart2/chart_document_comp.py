from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.style.style_families_supplier_partial import StyleFamiliesSupplierPartial
from ooodev.adapter.util.number_formats_supplier_partial import NumberFormatsSupplierPartial
from ooodev.adapter.chart2.chart_document_partial import ChartDocumentPartial
from ooodev.adapter.chart2.data.data_receiver_partial import DataReceiverPartial
from ooodev.adapter.chart2.titled_partial import TitledPartial

if TYPE_CHECKING:
    from com.sun.star.chart2 import ChartDocument  # service


class ChartDocumentComp(
    ComponentBase,
    ChartDocumentPartial,
    DataReceiverPartial,
    TitledPartial,
    StyleFamiliesSupplierPartial,
    NumberFormatsSupplierPartial,
):
    """
    Class for managing Chart2 ChartType Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: ChartDocument) -> None:
        """
        Constructor

        Args:
            component (ChartType): UNO Chart2 ChartType Component.
        """
        ComponentBase.__init__(self, component)
        ChartDocumentPartial.__init__(self, component=component, interface=None)
        DataReceiverPartial.__init__(self, component=component, interface=None)
        TitledPartial.__init__(self, component=component, interface=None)
        StyleFamiliesSupplierPartial.__init__(self, component=component, interface=None)
        NumberFormatsSupplierPartial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.chart2.ChartDocument",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> ChartDocument:
        """ChartDocument Component"""
        # pylint: disable=no-member
        return cast("ChartDocument", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
