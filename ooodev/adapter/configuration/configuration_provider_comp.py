from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

from com.sun.star.lang import XMultiServiceFactory

from ooodev.adapter._helper.builder import builder_helper
from ooodev.adapter._helper.builder.comp_defaults_partial import CompDefaultsPartial
from ooodev.adapter.component_prop import ComponentProp
from ooodev.adapter.lang.multi_service_factory_partial import MultiServiceFactoryPartial
from ooodev.adapter.lang.component_partial import ComponentPartial
from ooodev.utils.builder.default_builder import DefaultBuilder

if TYPE_CHECKING:
    from com.sun.star.configuration import ConfigurationProvider  # service
    from ooodev.loader.inst.lo_inst import LoInst


class _ConfigurationProviderComp(ComponentProp):

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
            "com.sun.star.configuration.ConfigurationProvider",
            "com.sun.star.configuration.DefaultProvider",
        )

    # region Properties
    @property
    def __class__(self):  # type: ignore
        # pretend to be a ConfigurationProviderComp class
        return ConfigurationProviderComp

    # endregion Properties


class ConfigurationProviderComp(
    _ConfigurationProviderComp, ComponentPartial, MultiServiceFactoryPartial, CompDefaultsPartial
):
    """
    Class for managing table ConfigurationProvider Component.

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
            name="ooodev.adapter.configuration.configuration_provider_comp.ConfigurationProviderComp",
            base_class=_ConfigurationProviderComp,
        )
        return inst

    # pylint: disable=unused-argument

    def __init__(self, component: ConfigurationProvider) -> None:
        """
        Constructor

        Args:
            component (ConfigurationProvider): UNO ConfigurationProvider Component.
        """
        # this it not actually called as __new__ is overridden
        pass

    @classmethod
    def from_lo(cls, lo_inst: LoInst | None = None, *args: Any) -> ConfigurationProviderComp:
        """
        Get the singleton instance from the Lo.

        Args:
            lo_inst (LoInst, optional): LoInst, Defaults to ``Lo.current_lo``.
            args (Any, optional): One or more args to pass to instance creation.

        Returns:
            ConfigurationProviderComp: The instance.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.loader import lo as mLo

        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        inst = lo_inst.create_instance_mcf(
            XMultiServiceFactory, "com.sun.star.configuration.ConfigurationProvider", args=args, raise_err=True
        )
        return cls(inst)  # type: ignore

    # region Properties
    @property
    @override
    def component(self) -> ConfigurationProvider:
        """ConfigurationProvider Component"""
        # pylint: disable=no-member
        return cast("ConfigurationProvider", self._ComponentBase__get_component())  # type: ignore

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
    builder.auto_add_interface("com.sun.star.lang.XMultiServiceFactory")
    builder.auto_add_interface("com.sun.star.lang.XComponent")
    return builder
