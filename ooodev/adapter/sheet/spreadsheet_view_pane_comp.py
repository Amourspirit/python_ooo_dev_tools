from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.sheet.cell_range_referrer_partial import CellRangeReferrerPartial
from ooodev.adapter.sheet.view_pane_partial import ViewPanePartial


if TYPE_CHECKING:
    from com.sun.star.sheet import SpreadsheetViewPane  # service


class SpreadsheetViewPaneComp(ComponentBase, ViewPanePartial, CellRangeReferrerPartial):
    """
    Class for managing SpreadsheetViewPane Component.

    .. versionadded:: 0.20.0
    """

    # Some implementations may also implement com.sun.star.view.XControlAccess

    def __init__(self, component: SpreadsheetViewPane) -> None:
        """
        Constructor

        Args:
            component (SpreadsheetViewPane): UNO Volatile Result Component
        """
        ComponentBase.__init__(self, component)
        ViewPanePartial.__init__(self, component=self.component, interface=None)
        CellRangeReferrerPartial.__init__(self, component=self.component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.sheet.SpreadsheetViewPane",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> SpreadsheetViewPane:
        """Volatile Result Component"""
        # pylint: disable=no-member
        return cast("SpreadsheetViewPane", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
