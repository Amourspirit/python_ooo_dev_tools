from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.container.enumeration_access_partial import EnumerationAccessPartial
from ooodev.adapter.container.index_access_partial import IndexAccessPartial
from ooodev.adapter.sheet.spreadsheets_partial import SpreadsheetsPartial

if TYPE_CHECKING:
    from com.sun.star.sheet import Spreadsheets  # service
    from com.sun.star.sheet import Spreadsheet  # service


class SpreadsheetsComp(
    ComponentBase, SpreadsheetsPartial, IndexAccessPartial["Spreadsheet"], EnumerationAccessPartial["Spreadsheet"]
):
    """
    Class for managing Spreadsheet Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Spreadsheets) -> None:
        """
        Constructor

        Args:
            component (Spreadsheets): UNO Spreadsheet Component
        """
        ComponentBase.__init__(self, component)
        SpreadsheetsPartial.__init__(self, component=component, interface=None)
        IndexAccessPartial.__init__(self, component=component, interface=None)
        EnumerationAccessPartial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.sheet.Spreadsheets",)

    # endregion Overrides

    # region Properties
    @property
    def component(self) -> Spreadsheets:
        """Spreadsheets Component"""
        # pylint: disable=no-member
        return cast("Spreadsheets", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
