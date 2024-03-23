from __future__ import annotations
from typing import TYPE_CHECKING
import uno

from com.sun.star.reflection import XTypeDescriptionEnumeration

from ooodev.adapter.container.enumeration_partial import EnumerationPartial

if TYPE_CHECKING:
    from com.sun.star.reflection import XTypeDescription
    from ooodev.utils.type_var import UnoInterface


class TypeDescriptionEnumerationPartial(EnumerationPartial["XTypeDescription"]):
    """
    Partial class for XTypeDescriptionEnumeration.
    """

    def __init__(
        self,
        component: XTypeDescriptionEnumeration,
        interface: UnoInterface | None = XTypeDescriptionEnumeration,
    ) -> None:
        """
        Constructor

        Args:
            component (XTypeDescriptionEnumeration): UNO Component that implements ``com.sun.star.reflection.XTypeDescriptionEnumeration`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XTypeDescriptionEnumeration``.
        """
        EnumerationPartial.__init__(self, component=component, interface=interface)
        self.__component = component

    # region XTypeDescriptionEnumeration
    def next_type_description(self) -> XTypeDescription:
        """
        Returns the next element of the enumeration.

        Raises:
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
        """
        return self.__component.nextTypeDescription()

    # endregion XTypeDescriptionEnumeration
