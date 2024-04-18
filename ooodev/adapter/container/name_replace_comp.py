from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING, Generic, TypeVar

from ooodev.adapter._helper.builder import builder_helper
from ooodev.adapter._helper.builder.comp_defaults_partial import CompDefaultsPartial
from ooodev.adapter.component_prop import ComponentProp
from ooodev.utils.builder.default_builder import DefaultBuilder
from ooodev.adapter.container import name_replace_partial

if TYPE_CHECKING:
    from com.sun.star.container import XNameReplace

T = TypeVar("T")


class _NameReplaceComp(ComponentProp):
    def __eq__(self, other: Any) -> bool:
        if not isinstance(other, _NameReplaceComp):
            return False
        if self is other:
            return True
        if self.component is other.component:
            return True
        return self.component == other.component


class NameReplaceComp(_NameReplaceComp, name_replace_partial.NameReplacePartial[T], CompDefaultsPartial, Generic[T]):
    """
    Class for managing XNameContainer Component.

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
            name="ooodev.adapter.container.name_replace_comp.NameReplaceComp",
            base_class=_NameReplaceComp,
        )
        return inst

    def __init__(self, component: XNameReplace) -> None:
        """
        Constructor

        Args:
            component (XNameReplace): UNO Component that implements ``com.sun.star.container.XNameReplace``.
        """

        # this it not actually called as __new__ is overridden
        pass

    # region Properties

    @property
    def component(self) -> XNameReplace:
        """XNameReplace Component"""
        # pylint: disable=no-member
        return cast("XNameReplace", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties


def get_builder(component: Any) -> DefaultBuilder:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """
    return name_replace_partial.get_builder(component=component)
