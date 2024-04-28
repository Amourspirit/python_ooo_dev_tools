from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.reflection import XConstantTypeDescription

from ooodev.adapter.reflection.type_description_partial import TypeDescriptionPartial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class ConstantTypeDescriptionPartial(TypeDescriptionPartial):
    """
    Partial class for XConstantTypeDescription.
    """

    def __init__(
        self,
        component: XConstantTypeDescription,
        interface: UnoInterface | None = XConstantTypeDescription,
    ) -> None:
        """
        Constructor

        Args:
            component (XConstantTypeDescription): UNO Component that implements ``com.sun.star.reflection.XConstantTypeDescription`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XConstantTypeDescription``.
        """
        TypeDescriptionPartial.__init__(self, component=component, interface=interface)
        self.__component = component

    # region XConstantTypeDescription
    def get_constant_value(self) -> Any:
        """
        Following types are allowed for constants:
        """
        return self.__component.getConstantValue()

    # endregion XConstantTypeDescription
