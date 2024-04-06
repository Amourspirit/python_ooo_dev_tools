from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from ooodev.adapter import builder_helper
from ooodev.adapter.beans import exact_name_partial
from ooodev.adapter.beans import property_partial
from ooodev.adapter.component_prop import ComponentProp
from ooodev.adapter.lang.service_info_partial import ServiceInfoPartial
from ooodev.utils.builder.check_kind import CheckKind
from ooodev.utils.builder.default_builder import DefaultBuilder
from ooodev.utils.builder.init_kind import InitKind
from ooodev.utils.partial.interface_partial import InterfacePartial
from ooodev.utils.partial.qi_partial import QiPartial

if TYPE_CHECKING:
    from com.sun.star.configuration import ConfigurationUpdateAccess  # service

    # from ooodev.utils.builder.default_builder import DefaultBuilder
    from ooodev.utils.inst.lo.lo_inst import LoInst


class _ConfigurationUpdateAccessComp(ComponentProp):

    def __eq__(self, other: Any) -> bool:
        if not isinstance(other, _ConfigurationUpdateAccessComp):
            return False
        if self is other:
            return True
        if self.component is other.component:
            return True
        return self.component == other.component

    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.configuration.ConfigurationUpdateAccess", "com.sun.star.configuration.DefaultProvider")


class ConfigurationUpdateAccessComp(
    _ConfigurationUpdateAccessComp,
    exact_name_partial.ExactNamePartial,
    property_partial.PropertyPartial,
    InterfacePartial,
    ServiceInfoPartial,
    QiPartial,
    # child_partial.ChildPartial,
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
        builder_helper.builder_add_service_defaults(builder)

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
        super().__init__(component)

    # region Properties

    @property
    def component(self) -> ConfigurationUpdateAccess:
        """ConfigurationUpdateAccess Component"""
        # pylint: disable=no-member
        return cast("ConfigurationUpdateAccess", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties

    @classmethod
    def from_lo(cls, node_name: str, lo_inst: LoInst | None = None, *args: Any) -> ConfigurationAccessComp:
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

        # inst = lo_inst.create_instance_mcf(
        #     XMultiServiceFactory, "com.sun.star.configuration.ConfigurationProvider", args=args, raise_err=True
        # )
        # return cls(inst)  # type: ignore
        return ConfigurationUpdateAccessComp(inst)


def get_builder(component: Any) -> DefaultBuilder:
    builder = DefaultBuilder(component)

    builder.auto_add_interface("com.sun.star.beans.XExactName")
    builder.auto_add_interface("com.sun.star.beans.XHierarchicalPropertySet")
    builder.auto_add_interface("com.sun.star.beans.XMultiHierarchicalPropertySet")
    builder.auto_add_interface("com.sun.star.beans.XMultiPropertyStates")
    builder.auto_add_interface("com.sun.star.beans.XProperty")
    builder.auto_add_interface("com.sun.star.beans.XPropertySetInfo")
    builder.auto_add_interface("com.sun.star.beans.XPropertyState")
    builder.auto_add_interface("com.sun.star.beans.XPropertyWithState")
    builder.auto_add_interface("com.sun.star.configuration.XTemplateContainer")
    builder.auto_add_interface("com.sun.star.configuration.XTemplateInstance")
    builder.auto_add_interface("com.sun.star.container.XChild")
    builder.auto_add_interface("com.sun.star.container.XContainer")
    builder.auto_add_interface("com.sun.star.container.XHierarchicalName")
    builder.auto_add_interface("com.sun.star.container.XHierarchicalNameAccess")
    builder.auto_add_interface("com.sun.star.container.XNameAccess")
    builder.auto_add_interface("com.sun.star.container.XNamed")
    builder.auto_add_interface("com.sun.star.lang.XComponent")
    builder.auto_add_interface("com.sun.star.lang.XLocalizable")
    builder.auto_add_interface("com.sun.star.util.XChangesNotifier")
    builder.auto_add_interface("com.sun.star.util.XStringEscape")
    builder.auto_add_interface("com.sun.star.util.XChangesBatch")

    builder.add_event(
        module_name="ooodev.adapter.container.container_events",
        class_name="ContainerEvents",
        uno_name="com.sun.star.container.XContainer",
        optional=True,
    )
    builder.add_event(
        module_name="ooodev.adapter.util.changes_events",
        class_name="ChangesEvents",
        uno_name="com.sun.star.util.XChangesNotifier",
        optional=True,
    )

    builder.add_import(
        name="ooodev.adapter.beans.property_change_implement.PropertyChangeImplement",
        uno_name="com.sun.star.beans.XPropertySet",
        optional=True,
        init_kind=InitKind.COMPONENT,
        check_kind=CheckKind.INTERFACE,
    )
    builder.add_import(
        name="ooodev.adapter.beans.vetoable_change_implement.VetoableChangeImplement",
        uno_name="com.sun.star.beans.XPropertySet",
        optional=True,
        init_kind=InitKind.COMPONENT,
        check_kind=CheckKind.INTERFACE,
    )
    builder.add_import(
        name="ooodev.adapter.beans.properties_change_implement.PropertiesChangeImplement",
        uno_name="com.sun.star.beans.XMultiPropertySet",
        optional=True,
        init_kind=InitKind.COMPONENT,
        check_kind=CheckKind.INTERFACE,
    )

    return builder
