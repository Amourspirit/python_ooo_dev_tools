from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

from ooodev.adapter.chart.chart_data_change_event_events import ChartDataChangeEventEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.adapter.table.cell_comp import CellComp
from ooodev.adapter.container.named_partial import NamedPartial
from .text_content_comp import TextContentComp


if TYPE_CHECKING:
    from com.sun.star.text import TextTable  # service
    from com.sun.star.text import XTextTable


class TextTableComp(TextContentComp, ChartDataChangeEventEvents, NamedPartial):
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
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        ChartDataChangeEventEvents.__init__(
            self, trigger_args=generic_args, cb=self._on_chart_data_change_event_add_remove
        )
        NamedPartial.__init__(self, component=self.component)

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
    def auto_format(self, name: str) -> None:
        """
        Applies an AutoFormat to the cell range of the current context.

        Args:
            name (str): The name of the table style to apply.
        """
        self.component.autoFormat(name)

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

    def get_cell_names(self) -> tuple[str, ...]:
        """
        Returns the names of all cells in the table.

        Returns:
            tuple[str, ...]: The names of all cells in the table.
        """
        return self.component.getCellNames()

    # endregion Methods

    # region Properties
    if TYPE_CHECKING:

        @property
        def component(self) -> TextTable:
            """TextTable Component"""
            return cast("TextTable", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
