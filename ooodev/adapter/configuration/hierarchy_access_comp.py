from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

from ooodev.adapter._helper.builder import builder_helper
from ooodev.adapter._helper.builder.comp_defaults_partial import CompDefaultsPartial
from ooodev.adapter.beans import exact_name_partial
from ooodev.adapter.component_prop import ComponentProp
from ooodev.adapter.container import container_partial
from ooodev.adapter.container import hierarchical_name_access_partial
from ooodev.adapter.container import name_access_partial
from ooodev.utils.builder.default_builder import DefaultBuilder
from ooodev.adapter.container.container_events import ContainerEvents

if TYPE_CHECKING:
    from com.sun.star.configuration import HierarchyAccess  # service


class _HierarchyAccessComp(ComponentProp):

    def __eq__(self, other: Any) -> bool:
        if not isinstance(other, ComponentProp):
            return False
        if self is other:
            return True
        if self.component is other.component:
            return True
        return self.component == other.component

    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.configuration.HierarchyAccess",)

    # region Properties
    @property
    def __class__(self):
        # pretend to be a HierarchyAccessComp class
        return HierarchyAccessComp

    # endregion Properties


class HierarchyAccessComp(
    _HierarchyAccessComp,
    name_access_partial.NameAccessPartial,
    hierarchical_name_access_partial.HierarchicalNameAccessPartial,
    container_partial.ContainerPartial,
    exact_name_partial.ExactNamePartial,
    ContainerEvents,
    CompDefaultsPartial,
):
    """
    Class for managing HierarchyAccess Component.

    Note:
        This is a Dynamic class that is created at runtime.
        This means that the class is created at runtime and not defined in the source code.
        In addition, the class may be created with additional classes implemented.

        The Type hints for this class at design time may not be accurate.
        To check if a class implements a specific interface, use the ``isinstance`` function
        or :py:meth:`~.InterfacePartial.is_supported_interface` methods which is always available in this class.
    """

    # pylint: disable=unused-argument

    def __new__(cls, component: Any, *args, **kwargs):
        builder = get_builder(component=component)
        builder_helper.builder_add_comp_defaults(builder)
        builder_only = kwargs.get("_builder_only", False)
        if builder_only:
            # cast to prevent type checker error
            return cast(Any, builder)
        inst = builder.build_class(
            name="ooodev.adapter.configuration.hierarchy_access_comp.HierarchyAccessComp",
            base_class=_HierarchyAccessComp,
        )
        return inst

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (XNameAccess): UNO Component that implements ``com.sun.star.container.XNameAccess``.
        """
        # this it not actually called as __new__ is overridden
        pass

    # region Properties

    @property
    def component(self) -> HierarchyAccess:
        """HierarchyAccess Component"""
        # pylint: disable=no-member
        return cast("HierarchyAccess", self._ComponentBase__get_component())  # type: ignore

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
    builder.set_omit("com.sun.star.container.XElementAccess")
    builder.auto_add_interface("com.sun.star.container.XNameAccess")
    builder.auto_add_interface("com.sun.star.container.XHierarchicalNameAccess")
    builder.auto_add_interface("com.sun.star.container.XContainer")
    builder.auto_add_interface("com.sun.star.beans.XExactName")
    builder.auto_add_interface("com.sun.star.beans.XPropertySetInfo")  # optional
    builder.auto_add_interface("com.sun.star.beans.XPropertyState")  # optional
    builder.auto_add_interface("com.sun.star.beans.XMultiPropertyStates")  # optional

    builder.add_event(
        module_name="ooodev.adapter.container.container_events",
        class_name="ContainerEvents",
        uno_name="com.sun.star.container.XContainer",
        optional=True,
    )
    return builder
