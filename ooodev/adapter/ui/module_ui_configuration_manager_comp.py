from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import contextlib
import uno
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.ui.accelerator_configuration_partial import AcceleratorConfigurationPartial
from ooodev.adapter.ui.ui_configuration_events import UIConfigurationEvents
from ooodev.adapter.ui.module_ui_configuration_manager2_partial import ModuleUIConfigurationManager2Partial


if TYPE_CHECKING:
    from com.sun.star.ui import ModuleUIConfigurationManager


class ModuleUIConfigurationManagerComp(ComponentBase, ModuleUIConfigurationManager2Partial, UIConfigurationEvents):
    """
    Class for managing ModuleUIConfigurationManager Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: ModuleUIConfigurationManager) -> None:
        """
        Constructor

        Args:
            component (ModuleUIConfigurationManager): UNO ModuleUIConfigurationManager Component that supports ``com.sun.star.ui.ModuleUIConfigurationManager`` service.
        """
        # component is a struct
        ComponentBase.__init__(self, component)
        ModuleUIConfigurationManager2Partial.__init__(self, component=component, interface=None)
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
        return ("com.sun.star.ui.ModuleUIConfigurationManager",)

    # endregion Overrides

    # region UIConfigurationManagerPartial overrides

    def has_settings(self, resource_url: str) -> bool:
        """
        Determines if the settings of a user interface element is part the user interface configuration manager.
        """
        with contextlib.suppress(Exception):
            return self.component.hasSettings(resource_url)
        return False

    # endregion UIConfigurationManagerPartial overrides

    def create_default(self, module_short_name: str, module_identifier: str) -> None:
        """
        Provides a function to initialize a module user interface configuration manager instance.

        A module user interface configuration manager instance needs the following arguments as com.sun.star.beans.PropertyValue to be in a working state:

        - ``DefaultConfigStorage`` a reference to a ``com.sun.star.embed.Storage`` that contains the default module user interface configuration settings.
        - ``UserConfigStorage`` a reference to a ``com.sun.star.embed.Storage`` that contains the user-defined module user interface configuration settings.
        - ``ModuleIdentifier`` string that provides the module identifier.
        - ``UserRootCommit`` a reference to a ``com.sun.star.embed.XTransactedObject`` which represents the customizable root storage. Every implementation must use this reference to commit its changes also at the root storage.

        A non-initialized module user interface configuration manager cannot be used, it is treated as a read-only container.

        Raises:
            com.sun.star.configuration.CorruptedUIConfigurationException: ``CorruptedUIConfigurationException``
            com.sun.star.beans.UnknownPropertyException: ``UnknownPropertyException``
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        self.component.createDefault(module_short_name, module_identifier)

    @property
    def component(self) -> ModuleUIConfigurationManager:
        """ModuleUIConfigurationManager Component"""
        # pylint: disable=no-member
        return cast("ModuleUIConfigurationManager", self._ComponentBase__get_component())  # type: ignore
