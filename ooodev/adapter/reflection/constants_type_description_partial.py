from __future__ import annotations
from typing import TYPE_CHECKING, Tuple
import uno

from com.sun.star.reflection import XConstantsTypeDescription

from ooodev.adapter.reflection.type_description_partial import TypeDescriptionPartial
from ooodev.adapter.reflection.constant_type_description_comp import ConstantTypeDescriptionComp

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class ConstantsTypeDescriptionPartial(TypeDescriptionPartial):
    """
    Partial class for XConstantsTypeDescription.
    """

    def __init__(
        self,
        component: XConstantsTypeDescription,
        interface: UnoInterface | None = XConstantsTypeDescription,
    ) -> None:
        """
        Constructor

        Args:
            component (XConstantsTypeDescription): UNO Component that implements ``com.sun.star.reflection.XConstantsTypeDescription`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XConstantsTypeDescription``.
        """
        TypeDescriptionPartial.__init__(self, component=component, interface=interface)
        self.__component = component

    # region XConstantsTypeDescription
    def get_constants(self) -> Tuple[ConstantTypeDescriptionComp, ...]:
        """
        Returns the constants defined for this constants group.
        """
        return tuple(ConstantTypeDescriptionComp(const) for const in self.__component.getConstants())

    # endregion XConstantsTypeDescription
