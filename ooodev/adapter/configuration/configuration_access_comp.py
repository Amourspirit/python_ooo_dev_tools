from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from com.sun.star.lang import XMultiServiceFactory
from ooodev.utils import info as mInfo
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.beans import exact_name_partial
from ooodev.adapter.container import name_access_partial
from ooodev.adapter.container import hierarchical_name_access_partial
from ooodev.adapter.container import hierarchical_name_partial
from ooodev.adapter.container import named_partial
from ooodev.adapter.container import container_partial
from ooodev.adapter.beans import property_partial
from ooodev.adapter.beans import property_set_info_partial
from ooodev.adapter.beans import property_state_partial
from ooodev.adapter.beans import multi_property_states_partial
from ooodev.adapter.beans import property_with_state_partial
from ooodev.adapter.container import child_partial
from ooodev.adapter.configuration import hierarchy_element_comp
from ooodev.adapter.configuration import hierarchy_access_comp


if TYPE_CHECKING:
    from com.sun.star.configuration import ConfigurationAccess  # service
    from ooodev.utils.builder.default_builder import DefaultBuilder
    from ooodev.utils.inst.lo.lo_inst import LoInst


class ConfigurationAccessComp(
    ComponentBase,
    exact_name_partial.ExactNamePartial,
    property_set_info_partial.PropertySetInfoPartial,
    property_state_partial.PropertyStatePartial,
    multi_property_states_partial.MultiPropertyStatesPartial,
    property_with_state_partial.PropertyWithStatePartial,
    child_partial.ChildPartial,
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

        ComponentBase.__init__(self, component)
        exact_name_partial.ExactNamePartial.__init__(self, component=component, interface=None)
        property_set_info_partial.PropertySetInfoPartial.__init__(self, component=component, interface=None)
        property_state_partial.PropertyStatePartial.__init__(self, component=component, interface=None)
        multi_property_states_partial.MultiPropertyStatesPartial.__init__(self, component=component, interface=None)
        property_with_state_partial.PropertyWithStatePartial.__init__(self, component=component, interface=None)
        child_partial.ChildPartial.__init__(self, component=component, interface=None)

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

        inst = lo_inst.create_instance_mcf(
            XMultiServiceFactory, "com.sun.star.configuration.ConfigurationProvider", args=args, raise_err=True
        )
        # return cls(inst)  # type: ignore
        builder = cast("DefaultBuilder", get_builder(inst, lo_inst, local=True))

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

        return builder.build_class(name="ConfigurationAccess", base_class=ConfigurationAccessComp)


def get_builder(component: Any, lo_inst: Any = None, **kwargs) -> Any:
    # pylint: disable=import-outside-toplevel
    from ooodev.utils.builder.default_builder import DefaultBuilder

    builder = DefaultBuilder(component, lo_inst)
    # when local this modules class is added as the base class.
    # When not local this modules base class in not included but all of its import classer are.
    # see from_lo() above.
    local = kwargs.get("local", False)

    # region exclude local builders

    inc_hnp = cast(DefaultBuilder, hierarchical_name_partial.get_builder(component, lo_inst))
    inc_ha = cast(DefaultBuilder, hierarchical_name_access_partial.get_builder(component, lo_inst))
    inc_np = cast(DefaultBuilder, named_partial.get_builder(component, lo_inst))
    inc_na = cast(DefaultBuilder, name_access_partial.get_builder(component, lo_inst))
    inc_cp = cast(DefaultBuilder, container_partial.get_builder(component, lo_inst))
    inc_pp = cast(DefaultBuilder, property_partial.get_builder(component, lo_inst))
    builder.omits.update(inc_hnp.omits)
    builder.omits.update(inc_ha.omits)
    builder.omits.update(inc_np.omits)
    builder.omits.update(inc_na.omits)
    builder.omits.update(inc_cp.omits)
    builder.omits.update(inc_pp.omits)

    if local:
        builder.set_omit(*inc_hnp.get_import_names())
        builder.set_omit(*inc_ha.get_import_names())
        builder.set_omit(*inc_na.get_import_names())
        builder.set_omit(*inc_cp.get_import_names())
        builder.set_omit(*inc_np.get_import_names())
        builder.set_omit(*inc_pp.get_import_names())
    else:
        builder.add_from_instance(inc_hnp, make_optional=True)
        builder.add_from_instance(inc_ha, make_optional=True)
        builder.add_from_instance(inc_na, make_optional=True)
        builder.add_from_instance(inc_cp, make_optional=True)
        builder.add_from_instance(inc_np, make_optional=True)
        builder.add_from_instance(inc_pp, make_optional=True)

    # endregion exclude local builders

    # region exclude other builders

    ex_nap = cast(DefaultBuilder, exact_name_partial.get_builder(component, lo_inst))
    builder.set_omit(*ex_nap.get_import_names())
    ex_psp = cast(DefaultBuilder, property_set_info_partial.get_builder(component, lo_inst))
    builder.set_omit(*ex_psp.get_import_names())
    ex_pst = cast(DefaultBuilder, property_state_partial.get_builder(component, lo_inst))
    builder.set_omit(*ex_pst.get_import_names())
    ex_mps = cast(DefaultBuilder, multi_property_states_partial.get_builder(component, lo_inst))
    builder.set_omit(*ex_mps.get_import_names())
    ex_pws = cast(DefaultBuilder, property_with_state_partial.get_builder(component, lo_inst))
    builder.set_omit(*ex_pws.get_import_names())
    ex_cp = cast(DefaultBuilder, child_partial.get_builder(component, lo_inst))
    builder.set_omit(*ex_cp.get_import_names())
    # builder.set_omit("ooodev.utils.partial.interface_partial.InterfacePartial")
    # endregion exclude other builders

    hac_builder = hierarchy_access_comp.get_builder(component, lo_inst)
    builder.add_from_instance(hac_builder, make_optional=True)

    hec_builder = hierarchy_element_comp.get_builder(component, lo_inst)
    builder.add_from_instance(hec_builder, make_optional=True)

    if mInfo.Info.support_service(component, "com.sun.star.configuration.SetAccess"):
        # com.sun.star.configuration.SetAccess service includes...
        #  com.sun.star.configuration.HierarchyAccess service
        #  com.sun.star.configuration.SimpleSetAccess service
        #  com.sun.star.container.XContainer          interface
        from ooodev.adapter.configuration import set_access_comp

        builder.add_from_instance(set_access_comp.get_builder(component, lo_inst), make_optional=False)

    if mInfo.Info.support_service(component, "com.sun.star.configuration.GroupAccess"):
        from ooodev.adapter.configuration import group_access_comp

        builder.add_from_instance(group_access_comp.get_builder(component, lo_inst), make_optional=False)

    if mInfo.Info.support_service(component, "com.sun.star.configuration.AccessRootElement"):
        from ooodev.adapter.configuration import access_root_element_comp

        builder.add_from_instance(access_root_element_comp.get_builder(component, lo_inst), make_optional=False)

    if mInfo.Info.support_service(component, "com.sun.star.configuration.SetElement"):
        from ooodev.adapter.configuration import set_element_comp

        builder.add_from_instance(set_element_comp.get_builder(component, lo_inst), make_optional=False)

    if mInfo.Info.support_service(component, "com.sun.star.configuration.GroupElement"):
        from ooodev.adapter.configuration import group_element_comp

        builder.add_from_instance(group_element_comp.get_builder(component, lo_inst), make_optional=False)

    builder.auto_add_interface("com.sun.star.util.XRefreshable")

    # ooodev.adapter.util.refresh_events
    builder.add_event(
        module_name="ooodev.adapter.util.refresh_events",
        class_name="RefreshEvents",
        uno_name="com.sun.star.util.XRefreshable",
        optional=True,
        check_kind=2,
    )
    return builder
