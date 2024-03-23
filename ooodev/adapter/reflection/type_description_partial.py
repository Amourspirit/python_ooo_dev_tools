from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.reflection import XTypeDescription
from ooo.dyn.uno.type_class import TypeClass

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class TypeDescriptionPartial:
    """
    Partial class for XTypeDescription.
    """

    def __init__(
        self,
        component: XTypeDescription,
        interface: UnoInterface | None = XTypeDescription,
    ) -> None:
        """
        Constructor

        Args:
            component (XTypeDescription): UNO Component that implements ``com.sun.star.reflection.XTypeDescription`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XTypeDescription``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XTypeDescription
    def get_name(self) -> str:
        """
        Returns the fully qualified name of the UNOIDL entity.
        """
        return self.__component.getName()

    def get_type_class(self) -> TypeClass:
        """
        Returns the type class of the reflected UNOIDL entity.

        Returns:
            TypeClass: The type class of the reflected UNOIDL entity.
        Hint:
            - ``TypeClass`` is an enum and can be imported from ``ooo.dyn.uno.type_class``.
        """
        return self.__component.getTypeClass()  # type: ignore

    # endregion XTypeDescription
