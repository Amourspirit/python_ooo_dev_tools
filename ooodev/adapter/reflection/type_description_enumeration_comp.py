from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.reflection.type_description_enumeration_partial import TypeDescriptionEnumerationPartial

if TYPE_CHECKING:
    from com.sun.star.reflection import XTypeDescriptionEnumeration


class TypeDescriptionEnumerationComp(ComponentBase, TypeDescriptionEnumerationPartial):
    """
    Class for managing Components.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTypeDescriptionEnumeration) -> None:
        """
        Constructor

        Args:
            component (XTypeDescriptionEnumeration): UNO Component that implements ``com.sun.star.reflection.XTypeDescriptionEnumeration`` interface.
        """
        ComponentBase.__init__(self, component)
        TypeDescriptionEnumerationPartial.__init__(self, component=component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        # validated by TypeDescriptionEnumerationPartial
        return ()

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> XTypeDescriptionEnumeration:
        """XTypeDescriptionEnumeration Component"""
        # pylint: disable=no-member
        return cast("XTypeDescriptionEnumeration", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
