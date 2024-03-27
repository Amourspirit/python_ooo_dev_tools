from __future__ import annotations
from typing import cast, TYPE_CHECKING

from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.awt.control_container_partial import ControlContainerPartial

if TYPE_CHECKING:
    from com.sun.star.awt import XControlContainer


class ControlContainerComp(ComponentBase, ControlContainerPartial):
    """
    Class for managing XControlContainer Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XControlContainer) -> None:
        """
        Constructor

        Args:
            component (UnoControlDialog): UNO Component that implements ``com.sun.star.awt.XControlContainer`` interface.
        """

        ComponentBase.__init__(self, component)
        ControlContainerPartial.__init__(self, component=component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ()

    # endregion Overrides

    # region Properties

    @property
    def component(self) -> XControlContainer:
        """XControlContainer Component"""
        # pylint: disable=no-member
        return cast("XControlContainer", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
