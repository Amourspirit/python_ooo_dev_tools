from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.reflection import XTypeDescriptionEnumerationAccess
from ooo.dyn.reflection.type_description_search_depth import TypeDescriptionSearchDepth
from ooo.dyn.uno.type_class import TypeClass

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.adapter.reflection.type_description_enumeration_comp import TypeDescriptionEnumerationComp

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class TypeDescriptionEnumerationAccessPartial:
    """
    Partial class for XTypeDescriptionEnumerationAccess.
    """

    def __init__(
        self,
        component: XTypeDescriptionEnumerationAccess,
        interface: UnoInterface | None = XTypeDescriptionEnumerationAccess,
    ) -> None:
        """
        Constructor

        Args:
            component (XTypeDescriptionEnumerationAccess): UNO Component that implements ``com.sun.star.reflection.XTypeDescriptionEnumerationAccess`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XTypeDescriptionEnumerationAccess``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XTypeDescriptionEnumerationAccess
    def create_type_description_enumeration(
        self, module_name: str, depth: TypeDescriptionSearchDepth, *types: TypeClass
    ) -> TypeDescriptionEnumerationComp:
        """
        Creates an enumeration for type descriptions.

        An enumeration is always created for a UNOIDL module. The enumeration contents can be restricted by specifying type classes. Only types that match one of the supplied type classes will be part of the collection. Additionally, it is possible to specify the depth for the search within the underlying type description tree.

        Valid types classes are:

        The enumeration returns implementations of XTypeDescription. Following concrete UNOIDL parts represented by specialized interfaces derived from XTypeDescription can be returned by the enumerator:

        Args:
            module_name (str): Name of the UNOIDL module.
            depth (TypeDescriptionSearchDepth): Depth for the search within the underlying type description tree.
            types (TypeClass): One or more type classes. Restricts the contents of the enumeration.
                It will only contain type descriptions that match one of the supplied type classes.
                An empty sequence specifies that the enumeration shall contain all type descriptions.

        Raises:
            NoSuchTypeNameException: ``NoSuchTypeNameException``
            InvalidTypeNameException: ``InvalidTypeNameException``

        Returns:
            TypeDescriptionEnumerationComp: Enumeration for type descriptions. ``for x in enums`` loop can be used to iterate through the enumeration.

        Note:
            Valid types classes are:

            - ``ooo.dyn.uno.type_class.MODULE``
            - ``ooo.dyn.uno.type_class.INTERFACE``
            - ``ooo.dyn.uno.type_class.SERVICE``
            - ``ooo.dyn.uno.type_class.STRUCT``
            - ``ooo.dyn.uno.type_class.ENUM``
            - ``ooo.dyn.uno.type_class.EXCEPTION``
            - ``ooo.dyn.uno.type_class.TYPEDEF``
            - ``ooo.dyn.uno.type_class.CONSTANT``
            - ``ooo.dyn.uno.type_class.CONSTANTS``
            - ``ooo.dyn.uno.type_class.SINGLETON``

        Hint:
            - ``TypeClass`` is an enum and can be imported from ``ooo.dyn.uno.type_class``.
        """
        return TypeDescriptionEnumerationComp(self.__component.createTypeDescriptionEnumeration(module_name, types, depth))  # type: ignore

    # endregion XTypeDescriptionEnumerationAccess
