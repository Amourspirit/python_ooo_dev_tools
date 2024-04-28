from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

from ooodev.adapter._helper.builder import builder_helper
from ooodev.adapter._helper.builder.comp_defaults_partial import CompDefaultsPartial
from ooodev.adapter.component_prop import ComponentProp
from ooodev.adapter.configuration import configuration_provider_comp
from ooodev.adapter.lang.multi_service_factory_partial import MultiServiceFactoryPartial
from ooodev.utils.builder.default_builder import DefaultBuilder

if TYPE_CHECKING:
    from com.sun.star.configuration import theDefaultProvider  # singleton
    from ooodev.loader.inst.lo_inst import LoInst


class _TheDefaultProviderComp(ComponentProp):

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
        return (
            "com.sun.star.configuration.DefaultProvider",
            "com.sun.star.configuration.theDefaultProvider",
        )

    # region Properties
    @property
    def __class__(self):
        # pretend to be a TheDefaultProviderComp class
        return TheDefaultProviderComp

    # endregion Properties


class TheDefaultProviderComp(_TheDefaultProviderComp, MultiServiceFactoryPartial, CompDefaultsPartial):
    """
    Class for managing theDefaultProvider Component.

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
            name="ooodev.adapter.configuration.the_default_provider_comp.TheDefaultProviderComp",
            base_class=_TheDefaultProviderComp,
        )

        return inst

    def __init__(self, component: theDefaultProvider) -> None:
        """
        Constructor

        Args:
            component (theDefaultProvider): UNO Component that implements ``com.sun.star.configuration.theDefaultProvider`` service.
        """
        # this it not actually called as __new__ is overridden
        pass

    @classmethod
    def from_lo(cls, lo_inst: LoInst | None = None) -> TheDefaultProviderComp:
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
        key = "com.sun.star.configuration.theDefaultProvider"
        if key in lo_inst.cache:
            return cast(TheDefaultProviderComp, lo_inst.cache[key])
        factory = lo_inst.get_singleton("/singletons/com.sun.star.configuration.theDefaultProvider")  # type: ignore
        if factory is None:
            raise ValueError("Could not get theDefaultProvider singleton.")
        inst = cls(factory)
        lo_inst.cache[key] = inst
        return inst

    # region Properties
    @property
    def component(self) -> theDefaultProvider:
        """theDefaultProvider Component"""
        # pylint: disable=no-member
        return cast("theDefaultProvider", self._ComponentBase__get_component())  # type: ignore

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
    builder.merge(configuration_provider_comp.get_builder(component), make_optional=True)
    builder.auto_add_interface("com.sun.star.util.XRefreshable")
    builder.auto_add_interface("com.sun.star.util.XFlushable")
    builder.auto_add_interface("com.sun.star.lang.XLocalizable")

    return builder
