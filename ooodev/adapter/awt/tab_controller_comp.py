from __future__ import annotations
from typing import cast, TYPE_CHECKING

from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.awt.tab_controller_partial import TabControllerPartial

if TYPE_CHECKING:
    from com.sun.star.awt import XTabController


class TabControllerComp(ComponentBase, TabControllerPartial):
    """
    Class for managing XTabController Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTabController) -> None:
        """
        Constructor

        Args:
            component (UnoControlDialog): UNO Component that implements ``com.sun.star.awt.XTabController`` interface.
        """

        ComponentBase.__init__(self, component)
        TabControllerPartial.__init__(self, component=component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ()

    # endregion Overrides

    # region Properties

    @property
    def component(self) -> XTabController:
        """XTabController Component"""
        # pylint: disable=no-member
        return cast("XTabController", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
