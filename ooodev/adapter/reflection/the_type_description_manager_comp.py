from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.container.set_partial import SetPartial
from ooodev.adapter.container.hierarchical_name_access_partial import HierarchicalNameAccessPartial
from ooodev.adapter.lang.component_partial import ComponentPartial
from ooodev.adapter.reflection.type_description_enumeration_access_partial import (
    TypeDescriptionEnumerationAccessPartial,
)


if TYPE_CHECKING:
    from com.sun.star.reflection import TypeDescriptionManager  # service


class TheTypeDescriptionManagerComp(
    ComponentBase,
    ComponentPartial,
    HierarchicalNameAccessPartial[Any],
    SetPartial,
    TypeDescriptionEnumerationAccessPartial,
):
    """
    Class for managing theGlobalEventBroadcaster Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: TypeDescriptionManager) -> None:
        """
        Constructor

        Args:
            component (TypeDescriptionManager): UNO Component that implements ``com.sun.star.reflection.TypeDescriptionManager`` service.
        """
        ComponentBase.__init__(self, component)
        ComponentPartial.__init__(self, component=component, interface=None)
        HierarchicalNameAccessPartial.__init__(self, component=component, interface=None)
        SetPartial.__init__(self, component=component, interface=None)
        TypeDescriptionEnumerationAccessPartial.__init__(self, component=component, interface=None)
        # pylint: disable=no-member

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ()

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> TypeDescriptionManager:
        """TypeDescriptionManager Component"""
        # pylint: disable=no-member
        return cast("TypeDescriptionManager", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
