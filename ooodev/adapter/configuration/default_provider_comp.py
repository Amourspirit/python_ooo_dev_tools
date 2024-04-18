from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno

from ooodev.adapter._helper.builder import builder_helper
from ooodev.adapter._helper.builder.comp_defaults_partial import CompDefaultsPartial
from ooodev.adapter.component_prop import ComponentProp
from ooodev.adapter.configuration import configuration_provider_comp
from ooodev.utils.builder.default_builder import DefaultBuilder


if TYPE_CHECKING:
    from com.sun.star.configuration import DefaultProvider  # service

    # from ooodev.utils.builder.default_builder import DefaultBuilder
    from ooodev.utils.inst.lo.lo_inst import LoInst


class _DefaultProviderComp(ComponentProp):

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
        return ("com.sun.star.configuration.DefaultProvider",)

    # region Properties
    @property
    def __class__(self):
        # pretend to be a DefaultProviderComp class
        return DefaultProviderComp

    # endregion Properties


class DefaultProviderComp(
    _DefaultProviderComp,
    configuration_provider_comp.ConfigurationProviderComp,
    CompDefaultsPartial,
):
    """
    Class for managing DefaultProvider Component.

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
            name="ooodev.adapter.configuration.configuration_access_comp.ConfigurationAccessComp",
            base_class=_DefaultProviderComp,
        )

        return inst

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Component that supports ``com.sun.star.configuration.DefaultProvider`` service.
        """
        # this it not actually called as __new__ is overridden
        super().__init__(component)

    # region Properties

    @property
    def component(self) -> DefaultProvider:
        """DefaultProvider Component"""
        # pylint: disable=no-member
        return cast("DefaultProvider", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties

    @classmethod
    def from_lo(cls, lo_inst: LoInst | None = None) -> DefaultProviderComp:
        """
        Get the singleton instance from the Lo.

        Args:
            lo_inst (LoInst, optional): LoInst, Defaults to ``Lo.current_lo``.

        Returns:
            ThePopupMenuControllerFactoryComp: The instance.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.loader import lo as mLo

        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        factory = lo_inst.get_singleton("/singletons/com.sun.star.configuration.theDefaultProvider")  # type: ignore
        if factory is None:
            raise ValueError("Could not get theDefaultProvider singleton.")
        return cls(factory)


def get_builder(component: Any) -> DefaultBuilder:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """
    builder = DefaultBuilder(component)
    builder.merge(configuration_provider_comp.get_builder(component), make_optional=True)
    builder.auto_add_interface("com.sun.star.util.XRefreshable")
    builder.auto_add_interface("com.sun.star.util.XFlushable")
    builder.auto_add_interface("com.sun.star.lang.XLocalizable")

    return builder
