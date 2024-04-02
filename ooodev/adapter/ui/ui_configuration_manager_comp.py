from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.ui.ui_configuration_manager_partial import UIConfigurationManagerPartial

if TYPE_CHECKING:
    from com.sun.star.ui import XUIConfigurationManager


class UIConfigurationManagerComp(ComponentBase, UIConfigurationManagerPartial):
    """
    Class for managing ``XUIConfigurationManager`` Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XUIConfigurationManager) -> None:
        """
        Constructor

        Args:
            component (XUIConfigurationManager): UNO Component that implements ``com.sun.star.ui.XUIConfigurationManager`` interface.
        """
        # component is a struct
        ComponentBase.__init__(self, component)
        UIConfigurationManagerPartial.__init__(self, component=component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        # validated by mTextRangePartial.TextRangePartial
        return ()

    # endregion Overrides

    @property
    def component(self) -> XUIConfigurationManager:
        """XUIConfigurationManager Component"""
        # pylint: disable=no-member
        return cast("XUIConfigurationManager", self._ComponentBase__get_component())  # type: ignore
