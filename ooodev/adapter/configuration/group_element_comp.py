from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

from ooodev.adapter._helper.builder import builder_helper
from ooodev.adapter._helper.builder.comp_defaults_partial import CompDefaultsPartial
from ooodev.adapter.component_prop import ComponentProp
from ooodev.adapter.container import child_partial
from ooodev.adapter.configuration import hierarchy_element_comp
from ooodev.utils.builder.default_builder import DefaultBuilder

if TYPE_CHECKING:
    from com.sun.star.configuration import GroupElement  # service


class _GroupElementComp(ComponentProp):
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
        return ("com.sun.star.configuration.GroupElement",)

    # region Properties
    @property
    def __class__(self):
        # pretend to be a GroupElementComp class
        return GroupElementComp

    # endregion Properties


class GroupElementComp(
    _GroupElementComp,
    hierarchy_element_comp.HierarchyElementComp,
    child_partial.ChildPartial,
    CompDefaultsPartial,
):
    """
    Class for managing GroupElement Component.

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
            name="ooodev.adapter.configuration.group_element_comp.GroupElementComp",
            base_class=_GroupElementComp,
        )
        return inst

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Component that supports ``com.sun.star.configuration.GroupElement`` service.
        """
        # this it not actually called as __new__ is overridden
        pass

    # region Properties

    @property
    def component(self) -> GroupElement:
        """GroupElement Component"""
        # pylint: disable=no-member
        return cast("GroupElement", self._ComponentBase__get_component())  # type: ignore

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

    builder.merge(hierarchy_element_comp.get_builder(component), make_optional=True)

    builder.auto_add_interface("com.sun.star.container.XChild")

    return builder
