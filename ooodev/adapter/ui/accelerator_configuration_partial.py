from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple

from com.sun.star.ui import XAcceleratorConfiguration
from ooodev.adapter.ui.ui_configuration_partial import UIConfigurationPartial
from ooodev.adapter.ui.ui_configuration_persistence_partial import UIConfigurationPersistencePartial
from ooodev.adapter.ui.ui_configuration_storage_partial import UIConfigurationStoragePartial
from ooodev.utils.builder.check_kind import CheckKind
from ooodev.utils.builder.default_builder import DefaultBuilder

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.awt import KeyEvent
    from ooodev.utils.type_var import UnoInterface


class AcceleratorConfigurationPartial(
    UIConfigurationPersistencePartial, UIConfigurationStoragePartial, UIConfigurationPartial
):
    """
    Partial Class for XAcceleratorConfiguration.
    """

    def __init__(
        self, component: XAcceleratorConfiguration, interface: UnoInterface | None = XAcceleratorConfiguration
    ) -> None:
        """
        Constructor

        Args:
            component (XAcceleratorConfiguration): UNO Component that implements ``com.sun.star.ui.XAcceleratorConfiguration``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XAcceleratorConfiguration``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        UIConfigurationPersistencePartial.__init__(self, component=component, interface=None)
        UIConfigurationStoragePartial.__init__(self, component=component, interface=None)
        UIConfigurationPartial.__init__(self, component=component, interface=None)
        self.__component = component

    # region XAcceleratorConfiguration
    def get_all_key_events(self) -> Tuple[KeyEvent, ...]:
        """
        Gets the list of all key events, which are available at this configuration set.

        The key events are the ``primary keys`` of this configuration sets. Means: Commands are registered for key events.

        Such key event can be mapped to its bound command, using the method ``get_command_for_key_event()``.
        """
        return self.__component.getAllKeyEvents()

    def get_command_by_key_event(self, key_event: KeyEvent) -> str:
        """
        Gets the registered command for the specified key event.

        This function can be used to:

        Raises:
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
        """
        return self.__component.getCommandByKeyEvent(key_event)

    def get_key_events_by_command(self, cmd: str) -> Tuple[KeyEvent, ...]:
        """
        Optimized access to the relation ``command-key`` instead of ``key-command`` which is provided normally by this interface.

        It can be used to implement collision handling, if more than one key event match to the same command.
        The returned list contains all possible key events - and the outside code can select a possible one.
        Of course - mostly this list will contain only one key event.

        Raises:
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
        """
        return self.__component.getKeyEventsByCommand(cmd)

    def get_preferred_key_events_for_command_list(self, *cmd_args: str) -> Tuple[KeyEvent, ...]:
        """
        Optimized function to map a list of commands to a corresponding list of key events.

        It provides a fast mapping, which is e.g. needed by a menu or toolbar implementation.
        E.g. a sub menu is described by a list of commands - and the implementation of the menu must show the corresponding shortcuts.
        Iteration over all items of this configuration set can be very expensive.

        Instead to the method ``get_key_events_for_command()`` the returned list contains only one(!) key event bound to one(!) requested command.
        If more than one key event is bound to a command - a selection is done inside this method. This internal selection can't be influenced from outside.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
        """
        # getPreferredKeyEventsForCommandList() can return a result of (None,)
        results = self.__component.getPreferredKeyEventsForCommandList(cmd_args)
        if not results:
            return ()
        return tuple([val for val in results if val is not None])  # type: ignore

    def remove_command_from_all_key_events(self, cmd: str) -> None:
        """
        search for a key-command-binding inside this configuration set, where the specified command is used.

        If such binding could be located, the command will be removed from it.
        If as result of that the key binding will be empty, if will be removed too.

        This is an optimized method, which can perform removing of commands from this configuration set.
        Because normally Commands are ``foreign keys`` and key identifier the ``primary keys`` - it needs some work to remove all commands outside this container.

        Raises:
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
        """
        self.__component.removeCommandFromAllKeyEvents(cmd)

    def remove_key_event(self, key_event: KeyEvent) -> None:
        """
        Remove a key-command-binding from this configuration set.

        Raises:
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
        """
        self.__component.removeKeyEvent(key_event)

    def set_key_event(self, key_event: KeyEvent, cmd: str) -> None:
        """
        Modify or create a key - command - binding.

        If the specified key event does not already exists inside this configuration access, it will be created and the command will be registered for it.
        If the specified key event already exists, its command will be overwritten with the new command.
        There is no warning nor any error about that! The outside code has to use the method getCommandForKeyEvent() to check for possible collisions.

        Note:
            This method can't be used to remove entities from the configuration set.
            Empty parameters will result into an exception! Use the method ``remove_key_event()`` instead.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
        """
        self.__component.setKeyEvent(key_event, cmd)

    # endregion XAcceleratorConfiguration


def get_builder(component: Any) -> DefaultBuilder:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """
    builder = DefaultBuilder(component)

    # for some unknown reason is not implemented in some components as a type but the methods are there.
    # Just add the interface, overriding the optional, if the store method is available.
    if hasattr(component, "getAllKeyEvents"):
        builder.auto_add_interface(
            "com.sun.star.ui.XAcceleratorConfiguration", optional=False, check_kind=CheckKind.NONE
        )

    if hasattr(component, "store"):
        builder.auto_add_interface(
            "com.sun.star.ui.XUIConfigurationPersistence", optional=False, check_kind=CheckKind.NONE
        )
    if hasattr(component, "hasStorage"):
        builder.auto_add_interface(
            "com.sun.star.ui.XUIConfigurationStorage", optional=False, check_kind=CheckKind.NONE
        )
    if hasattr(component, "addConfigurationListener"):
        builder.auto_add_interface("com.sun.star.ui.XUIConfiguration", optional=False, check_kind=CheckKind.NONE)
    return builder
