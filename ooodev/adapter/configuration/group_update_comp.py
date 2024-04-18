from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from ooodev.adapter._helper.builder import builder_helper
from ooodev.adapter._helper.builder.comp_defaults_partial import CompDefaultsPartial
from ooodev.adapter.component_prop import ComponentProp
from ooodev.utils.builder.default_builder import DefaultBuilder
from ooodev.adapter.configuration import group_access_comp
from ooodev.adapter.container import name_replace_partial


if TYPE_CHECKING:
    from com.sun.star.configuration import GroupUpdate  # service


class _GroupUpdateComp(ComponentProp):

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
        return ("com.sun.star.configuration.GroupUpdate",)

    # region Properties
    @property
    def __class__(self):
        # pretend to be a GroupUpdateComp class
        return GroupUpdateComp

    # endregion Properties


class GroupUpdateComp(
    _GroupUpdateComp,
    group_access_comp.GroupAccessComp,
    name_replace_partial.NameReplacePartial,
    CompDefaultsPartial,
):
    """
    Class for managing GroupUpdate Component.

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
            name="ooodev.adapter.configuration.group_access_comp.GroupAccessComp",
            base_class=_GroupUpdateComp,
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
    def component(self) -> GroupUpdate:
        """GroupUpdate Component"""
        # pylint: disable=no-member
        return cast("GroupUpdate", self._ComponentBase__get_component())  # type: ignore

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
    builder.merge(group_access_comp.get_builder(component=component), make_optional=True)
    builder.merge(name_replace_partial.get_builder(component=component), make_optional=True)

    return builder
