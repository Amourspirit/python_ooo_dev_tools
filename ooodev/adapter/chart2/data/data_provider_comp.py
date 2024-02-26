from __future__ import annotations
from typing import cast, TYPE_CHECKING
import contextlib

from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.chart2.data.data_provider_partial import DataProviderPartial

if TYPE_CHECKING:
    from com.sun.star.chart2.data import DataProvider  # service


class DataProviderComp(ComponentBase, DataProviderPartial):
    """
    Class for managing Chart2 Data Provider Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: DataProvider) -> None:
        """
        Constructor

        Args:
            component (DataProvider): UNO component that implements ``com.sun.star.chart2.data.DataProvider`` service.
        """
        ComponentBase.__init__(self, component)
        DataProviderPartial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        # validated by DataSourcePartial
        return ("com.sun.star.chart2.data.DataProvider",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> DataProvider:
        """DataProvider Component"""
        # pylint: disable=no-member
        return cast("DataProvider", self._ComponentBase__get_component())  # type: ignore

    @property
    def include_hidden_cells(self) -> bool | None:
        """
        Gets/Sets If set to false FALSE, values from hidden cells are not returned.

        **optional**.
        """
        with contextlib.suppress(AttributeError):
            return self.component.IncludeHiddenCells
        return None

    @include_hidden_cells.setter
    def include_hidden_cells(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.component.IncludeHiddenCells = value

    # endregion Properties
