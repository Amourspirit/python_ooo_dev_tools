from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

from ooodev.adapter._helper.builder import builder_helper
from ooodev.adapter._helper.builder.comp_defaults_partial import CompDefaultsPartial
from ooodev.adapter.component_prop import ComponentProp
from ooodev.adapter.beans.property_state_partial import PropertyStatePartial
from ooodev.adapter.beans.multi_property_states_partial import MultiPropertyStatesPartial
from ooodev.adapter.configuration import property_hierarchy_comp
from ooodev.adapter.configuration import hierarchy_access_comp
from ooodev.utils.builder.default_builder import DefaultBuilder


if TYPE_CHECKING:
    from com.sun.star.configuration import GroupAccess  # service


class _GroupAccessComp(ComponentProp):

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
        return ("com.sun.star.configuration.GroupAccess",)

    # region Properties
    @property
    def __class__(self):  # type: ignore
        # pretend to be a GroupAccessComp class
        return GroupAccessComp

    # endregion Properties


class GroupAccessComp(
    _GroupAccessComp,
    PropertyStatePartial,
    MultiPropertyStatesPartial,
    CompDefaultsPartial,
):
    """
    Class for managing GroupAccess Component.

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
            base_class=_GroupAccessComp,
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
    @override
    def component(self) -> GroupAccess:
        """GroupAccess Component"""
        # pylint: disable=no-member
        return cast("GroupAccess", self._ComponentBase__get_component())  # type: ignore

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
    builder.merge(property_hierarchy_comp.get_builder(component=component), make_optional=True)
    builder.merge(hierarchy_access_comp.get_builder(component=component), make_optional=True)
    builder.auto_add_interface("com.sun.star.beans.XMultiPropertyStates", optional=True)
    builder.auto_add_interface("com.sun.star.beans.XPropertyState", optional=True)

    return builder
