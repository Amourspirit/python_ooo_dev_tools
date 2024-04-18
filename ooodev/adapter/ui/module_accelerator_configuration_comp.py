from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.ui.accelerator_configuration_partial import AcceleratorConfigurationPartial
from ooodev.adapter.ui.ui_configuration_events import UIConfigurationEvents

if TYPE_CHECKING:
    from com.sun.star.ui import ModuleAcceleratorConfiguration


class ModuleAcceleratorConfigurationComp(ComponentBase, AcceleratorConfigurationPartial, UIConfigurationEvents):
    """
    Class for managing ModuleAcceleratorConfiguration Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: ModuleAcceleratorConfiguration) -> None:
        """
        Constructor

        Args:
            component (ModuleAcceleratorConfiguration): UNO ModuleAcceleratorConfiguration Component that supports ``com.sun.star.ui.ModuleAcceleratorConfiguration`` service.
        """
        # component is a struct
        ComponentBase.__init__(self, component)
        AcceleratorConfigurationPartial.__init__(self, component=component, interface=None)
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
        return ("com.sun.star.ui.ModuleAcceleratorConfiguration",)

    # endregion Overrides

    def create_with_module_identifier(self, module_identifier: str) -> None:
        """
        Creates the component with the specified module identifier.
        """
        self.component.createWithModuleIdentifier(module_identifier)

    @property
    def component(self) -> ModuleAcceleratorConfiguration:
        """ModuleAcceleratorConfiguration Component"""
        # pylint: disable=no-member
        return cast("ModuleAcceleratorConfiguration", self._ComponentBase__get_component())  # type: ignore
