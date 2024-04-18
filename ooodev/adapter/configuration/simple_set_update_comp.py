from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING, Tuple
import uno
from ooodev.adapter._helper.builder import builder_helper
from ooodev.adapter.component_prop import ComponentProp
from ooodev.utils.builder.default_builder import DefaultBuilder
from ooodev.adapter.configuration import simple_set_access_comp
from ooodev.adapter.container import name_container_partial
from ooodev.adapter.lang import combined_service_factory_partial
from ooodev.utils.builder.check_kind import CheckKind
from ooodev.utils.builder.init_kind import InitKind

if TYPE_CHECKING:
    from com.sun.star.configuration import SimpleSetUpdate  # service


class _SimpleSetUpdateComp(ComponentProp):

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
        return ("com.sun.star.configuration.SimpleSetUpdate",)

    # region Properties
    @property
    def __class__(self):
        # pretend to be a SimpleSetUpdateComp class
        return SimpleSetUpdateComp

    # endregion Properties


class SimpleSetUpdateComp(
    _SimpleSetUpdateComp, simple_set_access_comp.SimpleSetAccessComp, name_container_partial.NameContainerPartial
):
    """
    Class for managing SimpleSetUpdateComp Component.

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
            name="ooodev.adapter.configuration.simple_set_update_comp.SimpleSetUpdateComp",
            base_class=_SimpleSetUpdateComp,
        )
        return inst

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Component that supports ``com.sun.star.configuration.SimpleSetUpdate`` service.
        """

        # this it not actually called as __new__ is overridden
        pass

    # endregion Overrides

    # region Properties

    @property
    def component(self) -> SimpleSetUpdate:
        """SimpleSetUpdate Component"""
        # pylint: disable=no-member
        return cast("SimpleSetUpdate", self._ComponentBase__get_component())  # type: ignore

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
    builder.merge(simple_set_access_comp.get_builder(component), make_optional=True)
    builder.merge(name_container_partial.get_builder(component), make_optional=True)
    # region Service Factories
    # if only XMultiServiceFactory is supported then use MultiServiceFactoryPartial will be used
    # if only XSingleServiceFactory is supported then use SingleServiceFactoryPartial will be used
    # if both are supported then use CombinedServiceFactoryPartial will be used.
    builder.merge(combined_service_factory_partial.get_builder(component), make_optional=True)
    builder.add_import(
        name="ooodev.adapter.lang.multi_service_factory_partial.MultiServiceFactoryPartial",
        uno_name=cast(
            Tuple[str], ("com.sun.star.lang.XMultiServiceFactory", "com.sun.star.lang.XSingleServiceFactory")
        ),
        init_kind=InitKind.COMPONENT_INTERFACE,
        check_kind=CheckKind.INTERFACE_ONLY,
    )
    builder.add_import(
        name="ooodev.adapter.lang.single_service_factory_partial.SingleServiceFactoryPartial",
        uno_name=cast(
            Tuple[str], ("com.sun.star.lang.XSingleServiceFactory", "com.sun.star.lang.XMultiServiceFactory")
        ),
        init_kind=InitKind.COMPONENT_INTERFACE,
        check_kind=CheckKind.INTERFACE_ONLY,
    )
    # endregion Service Factories
    return builder
