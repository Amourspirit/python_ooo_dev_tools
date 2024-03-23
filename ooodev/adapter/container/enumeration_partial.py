from __future__ import annotations
from typing import Any, TYPE_CHECKING, Generic, TypeVar
import uno

from com.sun.star.container import XEnumeration

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface

T = TypeVar("T")


class EnumerationPartial(Generic[T]):
    """
    Partial class for XEnumeration.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XEnumeration, interface: UnoInterface | None = XEnumeration) -> None:
        """
        Constructor

        Args:
            component (XEnumeration): UNO Component that implements ``com.sun.star.container.XEnumeration`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XEnumeration``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    def __iter__(self):
        return self

    def __next__(self) -> T:
        if not self.__component.hasMoreElements():
            raise StopIteration

        while self.__component.hasMoreElements():
            return self.__component.nextElement()
        raise StopIteration

    # region XEnumeration
    def has_more_elements(self) -> bool:
        """
        tests whether this enumeration contains more elements.
        """
        return self.__component.hasMoreElements()

    def next_element(self) -> T:
        """
        Gets the next element of this enumeration.

        Returns:
            Any: the next element of this enumeration.
        """
        return self.__component.nextElement()

    # endregion XEnumeration
