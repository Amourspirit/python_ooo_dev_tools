from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.util import XChangesNotifier
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
    from com.sun.star.util import XChangesListener


class ChangesNotifierPartial:
    """
    Partial Class XChangesNotifier.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XChangesNotifier, interface: UnoInterface | None = XChangesNotifier) -> None:
        """
        Constructor

        Args:
            component (XChangesNotifier): UNO Component that implements ``com.sun.star.util.XChangesNotifier`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XChangesNotifier``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XChangesNotifier
    def add_changes_listener(self, listener: XChangesListener) -> None:
        """
        Adds the specified listener to receive events when changes occurred.
        """
        self.__component.addChangesListener(listener)

    def remove_changes_listener(self, listener: XChangesListener) -> None:
        """
        Removes the specified listener.
        """
        self.__component.removeChangesListener(listener)

    # endregion XChangesNotifier


def get_builder(component: Any) -> Any:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """
    # pylint: disable=import-outside-toplevel
    from ooodev.utils.builder.default_builder import DefaultBuilder

    builder = DefaultBuilder(component)
    builder.auto_add_interface("com.sun.star.util.XChangesNotifier", False)
    return builder
