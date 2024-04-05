from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from ooodev.utils import info as mInfo
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.beans import exact_name_partial
from ooodev.adapter.container import name_access_partial
from ooodev.adapter.beans import property_partial
from ooodev.adapter.container import child_partial
from ooodev.adapter.configuration import hierarchy_element_comp
from ooodev.adapter.configuration import hierarchy_access_comp
from ooodev.adapter.component_prop import ComponentProp


if TYPE_CHECKING:
    from com.sun.star.configuration import ConfigurationAccess  # service
    from ooodev.utils.builder.default_builder import DefaultBuilder
    from ooodev.utils.inst.lo.lo_inst import LoInst


class ConfigurationAccessComp(
    ComponentBase,
    exact_name_partial.ExactNamePartial,
    property_partial.PropertyPartial,
    # child_partial.ChildPartial,
):
    """
    Class for managing ConfigurationAccess Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (XNameAccess): UNO Component that implements ``com.sun.star.container.XNameAccess``.
        """
        # event though the api shows that XPropertySetInfo, XPropertyState, XMultiPropertyStates and XPropertyWithState
        # are supported, this is not always the case. The are optionally included in the builder.
        ComponentBase.__init__(self, component)
        exact_name_partial.ExactNamePartial.__init__(self, component=component, interface=None)
        property_partial.PropertyPartial.__init__(self, component=component, interface=None)
        # child_partial.ChildPartial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.configuration.ConfigurationAccess", "com.sun.star.configuration.DefaultProvider")

    # endregion Overrides

    # region Properties

    @property
    def component(self) -> ConfigurationAccess:
        """ConfigurationAccess Component"""
        # pylint: disable=no-member
        return cast("ConfigurationAccess", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties

    @classmethod
    def from_lo(cls, node_name: str, lo_inst: LoInst | None = None, *args: Any) -> Any:
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
        inst = cp.create_instance_with_arguments("com.sun.star.configuration.ConfigurationAccess", *cfg_args)

        # inst = lo_inst.create_instance_mcf(
        #     XMultiServiceFactory, "com.sun.star.configuration.ConfigurationProvider", args=args, raise_err=True
        # )
        # return cls(inst)  # type: ignore
        builder = cast("DefaultBuilder", get_builder(inst, lo_inst))

        # remove HierarchyAccessComp and use a base class
        # builder.remove_import("ooodev.adapter.configuration.hierarchy_access_comp.HierarchyAccessComp")

        builder.add_import(
            name="ooodev.utils.partial.interface_partial.InterfacePartial",
            optional=False,
            check_kind=0,
            init_kind=0,
        )
        builder.add_import(
            name="ooodev.utils.partial.qi_partial.QiPartial",
            uno_name="com.sun.star.uno.XInterface",
            optional=True,
            init_kind=2,
        )
        builder.add_import(
            name="ooodev.utils.partial.service_partial.ServicePartial",
            uno_name="com.sun.star.lang.XServiceInfo",
            optional=True,
            init_kind=2,
        )

        return builder.build_class(name="ConfigurationAccessComp", base_class=ComponentProp)


def get_builder(component: Any, lo_inst: Any = None) -> Any:
    # pylint: disable=import-outside-toplevel
    from ooodev.utils.builder.default_builder import DefaultBuilder

    builder = DefaultBuilder(component, lo_inst)

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
        init_kind=1,
        check_kind=2,
    )
    builder.add_import(
        name="ooodev.adapter.beans.vetoable_change_implement.VetoableChangeImplement",
        uno_name="com.sun.star.beans.XPropertySet",
        optional=True,
        init_kind=1,
        check_kind=2,
    )
    builder.add_import(
        name="ooodev.adapter.beans.properties_change_implement.PropertiesChangeImplement",
        uno_name="com.sun.star.beans.XMultiPropertySet",
        optional=True,
        init_kind=1,
        check_kind=2,
    )

    return builder
