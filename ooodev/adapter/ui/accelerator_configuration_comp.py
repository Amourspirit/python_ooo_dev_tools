from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.ui.accelerator_configuration_partial import AcceleratorConfigurationPartial
from ooodev.adapter.ui.ui_configuration_events import UIConfigurationEvents

if TYPE_CHECKING:
    from com.sun.star.ui import XAcceleratorConfiguration


class AcceleratorConfigurationComp(ComponentBase, AcceleratorConfigurationPartial, UIConfigurationEvents):
    """
    Class for managing XAcceleratorConfiguration Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XAcceleratorConfiguration) -> None:
        """
        Constructor

        Args:
            component (XAcceleratorConfiguration): UNO Component that implements ``com.sun.star.ui.XAcceleratorConfiguration`` interface.
        """
        # component is a struct
        ComponentBase.__init__(self, component)
        AcceleratorConfigurationPartial.__init__(self, component=component)
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        UIConfigurationEvents.__init__(
            self, trigger_args=generic_args, cb=self.__on_ui_configuration_events_add_remove
        )

    # region Lazy Listeners
    def __on_ui_configuration_events_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addConfigurationListener(self.events_listener_ui_configuration)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        # validated by mTextRangePartial.TextRangePartial
        return ()

    # endregion Overrides

    @property
    def component(self) -> XAcceleratorConfiguration:
        """XAcceleratorConfiguration Component"""
        # pylint: disable=no-member
        return cast("XAcceleratorConfiguration", self._ComponentBase__get_component())  # type: ignore
