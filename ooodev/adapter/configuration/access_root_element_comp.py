from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

from ooodev.adapter.beans import property_with_state_partial
from ooodev.adapter._helper.builder.comp_defaults_partial import CompDefaultsPartial
from ooodev.adapter.component_prop import ComponentProp
from ooodev.adapter.configuration import hierarchy_element_comp
from ooodev.adapter.container import child_partial
from ooodev.adapter.lang import component_partial
from ooodev.adapter.util import changes_events
from ooodev.adapter.util import changes_notifier_partial
from ooodev.utils.builder.default_builder import DefaultBuilder

if TYPE_CHECKING:
    from com.sun.star.configuration import AccessRootElement  # service


class _AccessRootElementComp(hierarchy_element_comp._HierarchyElementComp):

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
        return ("com.sun.star.configuration.AccessRootElement",)

    # region Properties
    @property
    def __class__(self):
        # pretend to be a AccessRootElementComp class
        return AccessRootElementComp

    # endregion Properties


class AccessRootElementComp(
    _AccessRootElementComp,
    property_with_state_partial.PropertyWithStatePartial,
    child_partial.ChildPartial,
    component_partial.ComponentPartial,
    changes_notifier_partial.ChangesNotifierPartial,
    changes_events.ChangesEvents,
    CompDefaultsPartial,
):
    """
    Class for managing AccessRootElement Component.

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
        builder = cast(
            DefaultBuilder,
            hierarchy_element_comp.HierarchyElementComp.__new__(cls, component, _builder_only=True, *args, **kwargs),
        )

        local_builder = get_builder(component=component, _for_new=True)
        builder.merge(local_builder)

        builder_only = kwargs.get("_builder_only", False)
        if builder_only:
            # cast to prevent type checker error
            return cast(Any, builder)
        inst = builder.build_class(
            name="ooodev.adapter.configuration.access_root_element_comp.AccessRootElementComp",
            base_class=_AccessRootElementComp,
        )

        return inst

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Component that supports ``com.sun.star.configuration.AccessRootElement`` service.
        """
        # this it not actually called as __new__ is overridden
        pass

    # region Properties

    @property
    def component(self) -> AccessRootElement:
        """AccessRootElement Component"""
        # pylint: disable=no-member
        return cast("AccessRootElement", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties


def get_builder(component: Any, **kwargs) -> DefaultBuilder:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """
    for_new = kwargs.get("_for_new", False)
    if for_new:
        builder = DefaultBuilder(component)
    else:
        builder = hierarchy_element_comp.get_builder(component)

    # in a from_lo method in this class the HierarchyAccessComp would be removed and used as the base class
    builder.auto_add_interface("com.sun.star.lang.XComponent")
    builder.auto_add_interface("com.sun.star.util.XChangesNotifier")
    builder.auto_add_interface("com.sun.star.lang.XLocalizable")

    builder.add_event(
        module_name="ooodev.adapter.util.changes_events",
        class_name="ChangesEvents",
        uno_name="com.sun.star.util.XChangesNotifier",
        optional=True,
    )

    return builder
