from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.util import XFlushable

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.util import XFlushListener
    from ooodev.utils.type_var import UnoInterface


class FlushablePartial:
    """
    Partial Class XFlushable.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XFlushable, interface: UnoInterface | None = XFlushable) -> None:
        """
        Constructor

        Args:
            component (XFlushable): UNO Component that implements ``com.sun.star.util.XFlushable`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XFlushable``.
        """

        def validate(component: Any, interface: Any) -> None:
            if interface is None:
                return
            if not mLo.Lo.is_uno_interfaces(component, interface):
                raise mEx.MissingInterfaceError(interface)

        validate(component, interface)
        self.__component = component

    # region XFlushable
    def add_flush_listener(self, listener: XFlushListener) -> None:
        """
        Adds the specified listener to receive event ``flushed``.
        """
        self.__component.addFlushListener(listener)

    def flush(self) -> None:
        """
        Flushes the data of the object to the connected data source.
        """
        self.__component.flush()

    def remove_flush_listener(self, listener: XFlushListener) -> None:
        """
        Removes the specified listener.
        """
        self.__component.removeFlushListener(listener)

    # endregion XFlushable


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
    builder.auto_add_interface("com.sun.star.util.XFlushListener", False)
    return builder
