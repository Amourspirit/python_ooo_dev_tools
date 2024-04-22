from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

from ooodev.adapter.component_prop import ComponentProp
from ooodev.utils.builder.default_builder import DefaultBuilder
from ooodev.adapter.uno.weak_partial import WeakPartial
from ooodev.adapter.lang.component_partial import ComponentPartial
from ooodev.adapter.lang.type_provider_partial import TypeProviderPartial
from ooodev.adapter.lang.service_info_partial import ServiceInfoPartial
from ooodev.adapter.container.hierarchical_name_access_partial import HierarchicalNameAccessPartial
from ooodev.adapter.container import set_partial
from ooodev.adapter.reflection.type_description_enumeration_access_partial import (
    TypeDescriptionEnumerationAccessPartial,
)


if TYPE_CHECKING:
    from com.sun.star.reflection import TypeDescriptionManager
    from ooodev.loader.inst.lo_inst import LoInst

# theTypeDescriptionManager is not documented in the api.


class TheTypeDescriptionManagerComp(
    ComponentProp,
    WeakPartial,
    ComponentPartial,
    TypeProviderPartial,
    ServiceInfoPartial,
    HierarchicalNameAccessPartial[Any],
    set_partial.SetPartial,
    TypeDescriptionEnumerationAccessPartial,
):
    """
    Class for managing theTypeDescriptionManager Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: TypeDescriptionManager) -> None:
        """
        Constructor

        Args:
            component (TypeDescriptionManager): UNO Component that implements ``com.sun.star.reflection.TypeDescriptionManager`` service.
        """
        # this it not actually called as __new__ is overridden
        ComponentProp.__init__(self, component)
        WeakPartial.__init__(self, component=component, interface=None)  # type: ignore
        ComponentPartial.__init__(self, component, interface=None)  # type: ignore
        TypeProviderPartial.__init__(self, component, interface=None)  # type: ignore
        ServiceInfoPartial.__init__(self, component, interface=None)  # type: ignore
        HierarchicalNameAccessPartial.__init__(self, component, interface=None)
        set_partial.SetPartial.__init__(self, component, interface=None)
        TypeDescriptionEnumerationAccessPartial.__init__(self, component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.reflection.TypeDescriptionManager",)

    # endregion Overrides

    # region Class Methods
    @classmethod
    def from_lo(cls, lo_inst: LoInst | None = None) -> TheTypeDescriptionManagerComp:
        """
        Get the singleton instance from the Lo.

        Args:
            lo_inst (LoInst, optional): LoInst, Defaults to ``Lo.current_lo``.

        Returns:
            TheTypeDescriptionManagerComp: The instance.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.loader import lo as mLo

        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo

        key = "com.sun.star.reflection.theTypeDescriptionManager"
        if key in lo_inst.cache:
            return cast(TheTypeDescriptionManagerComp, lo_inst.cache[key])
        factory = lo_inst.get_singleton("/singletons/com.sun.star.reflection.theTypeDescriptionManager")  # type: ignore
        if factory is None:
            raise ValueError("Could not get theDefaultProvider singleton.")
        inst = cls(factory)
        lo_inst.cache[key] = inst
        return cast(TheTypeDescriptionManagerComp, inst)

    # endregion Class Methods

    # region Properties
    @property
    def component(self) -> TypeDescriptionManager:
        """TypeDescriptionManager Component"""
        # pylint: disable=no-member
        return cast("TypeDescriptionManager", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties


def get_builder(component: Any) -> DefaultBuilder:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """
    builder = DefaultBuilder(component)
    builder.auto_add_interface("com.sun.star.uno.XWeak")
    builder.auto_add_interface("com.sun.star.lang.XComponent")
    builder.auto_add_interface("com.sun.star.lang.XTypeProvider")
    builder.auto_add_interface("com.sun.star.lang.XServiceInfo")
    builder.auto_add_interface("com.sun.star.container.XHierarchicalNameAccess")
    builder.merge(set_partial.get_builder(component))
    builder.auto_add_interface("com.sun.star.reflection.XTypeDescriptionEnumerationAccess")

    return builder
