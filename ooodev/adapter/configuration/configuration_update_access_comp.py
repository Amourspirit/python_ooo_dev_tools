from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from ooodev.adapter._helper.builder import builder_helper
from ooodev.adapter.component_prop import ComponentProp
from ooodev.adapter._helper.builder.comp_defaults_partial import CompDefaultsPartial
from ooodev.utils.builder.default_builder import DefaultBuilder
from ooodev.adapter.configuration import configuration_access_comp
from ooodev.adapter.configuration import set_update_comp
from ooodev.adapter.configuration import group_update_comp
from ooodev.adapter.configuration import update_root_element_comp

if TYPE_CHECKING:
    from com.sun.star.configuration import ConfigurationUpdateAccess  # service

    # from ooodev.utils.builder.default_builder import DefaultBuilder
    from ooodev.utils.inst.lo.lo_inst import LoInst


class _ConfigurationUpdateAccessComp(ComponentProp):

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
        return ("com.sun.star.configuration.ConfigurationUpdateAccess",)

    # region Properties
    @property
    def __class__(self):
        # pretend to be a ConfigurationUpdateAccessComp class
        return ConfigurationUpdateAccessComp

    # endregion Properties


class ConfigurationUpdateAccessComp(
    _ConfigurationUpdateAccessComp,
    configuration_access_comp.ConfigurationAccessComp,
    CompDefaultsPartial,
):
    """
    Class for managing ConfigurationUpdateAccess Component.

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
            base_class=_ConfigurationUpdateAccessComp,
        )

        return inst

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Component that supports ``com.sun.star.configuration.ConfigurationUpdateAccess`` service.
        """
        # this it not actually called as __new__ is overridden
        pass

    # region Properties

    @property
    def component(self) -> ConfigurationUpdateAccess:
        """ConfigurationUpdateAccess Component"""
        # pylint: disable=no-member
        return cast("ConfigurationUpdateAccess", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties

    @classmethod
    def from_lo(cls, node_name: str, lo_inst: LoInst | None = None, *args: Any) -> ConfigurationUpdateAccessComp:
        """
        Get the singleton instance from the Lo.

        Args:
            lo_inst (LoInst, optional): LoInst, Defaults to ``Lo.current_lo``.
            args (Any, optional): One or more args to pass to instance creation.

        Returns:
            ConfigurationProviderComp: The instance with additional classes implemented.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.loader import lo as mLo
        from ooodev.adapter.configuration.configuration_provider_comp import ConfigurationProviderComp
        from ooodev.utils.props import Props

        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        cfg_args = list(args)
        cfg_args.append(Props.make_prop_value(name="nodepath", value=node_name))

        cp = ConfigurationProviderComp.from_lo(lo_inst=lo_inst)
        inst = cp.create_instance_with_arguments(
            service_name="com.sun.star.configuration.ConfigurationUpdateAccess", *cfg_args
        )

        return cls(inst)


def get_builder(component: Any) -> DefaultBuilder:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """
    builder = DefaultBuilder(component)
    builder.merge(configuration_access_comp.get_builder(component), make_optional=True)
    builder.merge(set_update_comp.get_builder(component), make_optional=True)
    builder.merge(group_update_comp.get_builder(component), make_optional=True)
    builder.merge(update_root_element_comp.get_builder(component), make_optional=True)

    return builder
