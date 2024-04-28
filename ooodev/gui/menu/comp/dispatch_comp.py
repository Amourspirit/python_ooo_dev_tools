from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from ooodev.adapter.component_prop import ComponentProp
from ooodev.utils.builder.default_builder import DefaultBuilder
from ooodev.adapter.frame.notifying_dispatch_partial import NotifyingDispatchPartial
from ooodev.adapter.lang.type_provider_partial import TypeProviderPartial
from ooodev.adapter.uno.weak_partial import WeakPartial

if TYPE_CHECKING:
    from com.sun.star.frame import XNotifyingDispatch


class _DispatchComp(ComponentProp):

    def __eq__(self, other: Any) -> bool:
        if not isinstance(other, ComponentProp):
            return False
        if self is other:
            return True
        if self.component is other.component:
            return True
        return self.component == other.component


class DispatchComp(
    _DispatchComp,
    NotifyingDispatchPartial,
    TypeProviderPartial,
    WeakPartial,
    # child_partial.ChildPartial,
):
    """
    Class for managing XNotifyingDispatch Component.

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
        # builder_helper.builder_add_comp_defaults(builder)

        builder_only = kwargs.get("_builder_only", False)
        if builder_only:
            # cast to prevent type checker error
            return cast(Any, builder)
        inst = builder.build_class(
            name="ooodev.gui.menu.comp.dispatch_comp.DispatchComp",
            base_class=_DispatchComp,
        )
        return inst

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Component that implements ``com.sun.star.frame.XNotifyingDispatch`` interface.
        """
        # this it not actually called as __new__ is overridden
        pass

    # region Properties

    @property
    def component(self) -> XNotifyingDispatch:
        """XNotifyingDispatch Component"""
        # pylint: disable=no-member
        return cast("XNotifyingDispatch", self._ComponentBase__get_component())  # type: ignore

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

    builder.auto_add_interface("com.sun.star.lang.XTypeProvider")
    builder.auto_add_interface("com.sun.star.frame.XNotifyingDispatch")
    builder.auto_add_interface("com.sun.star.uno.XWeak")

    return builder
