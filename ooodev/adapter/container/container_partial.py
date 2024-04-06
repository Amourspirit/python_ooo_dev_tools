from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.container import XContainer

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils.type_var import UnoInterface

if TYPE_CHECKING:
    from com.sun.star.container import XContainerListener


class ContainerPartial:
    """
    Partial class for XContainer.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XContainer, interface: UnoInterface | None = XContainer) -> None:
        """
        Constructor

        Args:
            component (XContainer): UNO Component that implements ``com.sun.star.container.XContainer`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XContainer``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XContainer
    def add_container_listener(self, listener: XContainerListener) -> None:
        """
        Adds the specified listener to receive events when elements are inserted or removed.

        Args:
            listener (XContainerListener): The listener to be added.
        """
        self.__component.addContainerListener(listener)

    def remove_container_listener(self, listener: XContainerListener) -> None:
        """
        Removes the specified listener so it does not receive any events from this container.

        Args:
            listener (XContainerListener): The listener to be removed.
        """
        self.__component.removeContainerListener(listener)

    # endregion XContainer


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
    builder.add_import(
        name="ooodev.adapter.container.container_partial.ContainerPartial",
        uno_name="com.sun.star.container.XContainer",
        optional=False,
        init_kind=2,
    )
    return builder
