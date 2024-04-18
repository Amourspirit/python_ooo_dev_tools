from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno

from com.sun.star.lang import XTypeProvider

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class TypeProviderPartial:
    """
    Partial class for ``XTypeProvider``.
    """

    def __init__(self, component: XTypeProvider, interface: UnoInterface | None = XTypeProvider) -> None:
        """
        Constructor

        Args:
            component (XTypeProvider): UNO Component that implements ``com.sun.star.lang.XTypeProvider`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XTypeProvider``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XTypeProvider
    def get_implementation_id(self) -> uno.ByteSequence:
        """
        Obsolete unique identifier.

        Originally returned a sequence of bytes which, when non-empty,
        was used as an ID to distinguish unambiguously between two sets of types,
        for example to realize hashing functionality when the object is introspected.
        Two objects that returned the same non-empty ID had to return the same set of types in getTypes().
        (If a unique ID could not be provided, this method was always allowed to return an empty sequence, though).
        """
        return self.__component.getImplementationId()

    def get_types(self) -> Tuple[Any, ...]:
        """
        returns a sequence of all types (usually interface types) provided by the object.
        """
        return self.__component.getTypes()

    # endregion XTypeProvider


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
    builder.auto_add_interface("com.sun.star.lang.XTypeProvider", False)
    return builder
