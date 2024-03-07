from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

from ooodev.adapter.chart.chart_data_change_event_events import ChartDataChangeEventEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.adapter.table.cell_comp import CellComp
from ooodev.adapter.container.named_partial import NamedPartial
from ooodev.adapter.text.text_content_comp import TextContentComp
from ooodev.adapter.text.text_table_partial import TextTablePartial
from ooodev.adapter.table.cell_range_partial import CellRangePartial
from ooodev.adapter.table.auto_formattable_partial import AutoFormattablePartial
from ooodev.adapter.sheet.cell_range_data_partial import CellRangeDataPartial
from ooodev.adapter.text.text_table_properties_partial import TextTablePropertiesPartial


if TYPE_CHECKING:
    from com.sun.star.text import TextTable  # service
    from com.sun.star.text import XTextTable


class TextTableComp(
    TextContentComp,
    TextTablePropertiesPartial,
    TextTablePartial,
    CellRangePartial,
    AutoFormattablePartial,
    CellRangeDataPartial,
    ChartDataChangeEventEvents,
    NamedPartial,
):
    """
    Class for managing TextTable Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTextTable) -> None:
        """
        Constructor

        Args:
            component (XTextTable): UNO Component that support ``com.sun.star.text.TextTable`` service.
        """

        TextContentComp.__init__(self, component)
        TextTablePropertiesPartial.__init__(self, component=component)  # type: ignore
        TextTablePartial.__init__(self, component=component, interface=None)
        CellRangePartial.__init__(self, component=component, interface=None)  # type: ignore
        AutoFormattablePartial.__init__(self, component=component, interface=None)  # type: ignore
        CellRangeDataPartial.__init__(self, component=component, interface=None)  # type: ignore
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        ChartDataChangeEventEvents.__init__(
            self, trigger_args=generic_args, cb=self._on_chart_data_change_event_add_remove
        )
        NamedPartial.__init__(self, component=self.component, interface=None)

    # region Lazy Listeners
    def _on_chart_data_change_event_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addChartDataChangeEventListener(self.events_listener_chart_data_change_event)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.text.TextTable",)

    # endregion Overrides

    # region Methods
    # region TextTablePartial Overrides
    def get_cell_by_name(self, name: str) -> CellComp:
        """
        Returns the cell with the specified name.

        The cell in the 4th column and third row has the name ``D3``.

        Args:
            name (str): The name of the cell.

        Returns:
            CellComp: The cell with the specified name.

        Note:
            In cells that are split, the naming convention is more complex.
            In this case the name is a concatenation of the former cell name (i.e. ``D3``) and
            the number of the new column and row index inside of the original table cell separated by dots.
            This is done recursively.

            For example, if the cell ``D3`` is horizontally split, it now contains the cells ``D3.1.1`` and ``D3.1.2``.
        """
        return CellComp(self.component.getCellByName(name))

    # endregion TextTablePartial Overrides

    # endregion Methods

    # region Properties
    if TYPE_CHECKING:

        @property
        def component(self) -> TextTable:
            """TextTable Component"""
            # pylint: disable=no-member
            return cast("TextTable", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
