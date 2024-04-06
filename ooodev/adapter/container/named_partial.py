from __future__ import annotations
from typing import Any
import uno

from com.sun.star.container import XNamed

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils.type_var import UnoInterface


class NamedPartial:
    """
    Partial class for XNamed.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XNamed, interface: UnoInterface | None = XNamed) -> None:
        """
        Constructor

        Args:
            component (XNamed): UNO Component that implements ``com.sun.star.container.XNamed`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XNamed``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XNamed
    def get_name(self) -> str:
        """Returns the name of the object."""
        return self.__component.getName()

    def set_name(self, name: str) -> None:
        """Sets the name of the object."""
        self.__component.setName(name)

    # endregion XNamed


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
        name="ooodev.adapter.container.named_partial.NamedPartial",
        uno_name="com.sun.star.container.XNamed",
        optional=False,
        init_kind=2,
    )
    return builder
