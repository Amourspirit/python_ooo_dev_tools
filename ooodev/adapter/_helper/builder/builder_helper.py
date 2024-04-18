from __future__ import annotations
from ooodev.utils.builder.default_builder import DefaultBuilder
from ooodev.utils.builder.check_kind import CheckKind
from ooodev.utils.builder.init_kind import InitKind


def builder_add_comp_defaults(builder: DefaultBuilder) -> None:
    """
    Adds the default imports for the builder.

    Args:
        builder (DefaultBuilder): A Builder instance.

    Note:
        Adds the following imports:

        - ``ServiceInfoPartial``
        - ``TypeProviderPartial``
        - ``InterfacePartial``
        - ``QiPartial``
        - ``EventsPartial``
    """
    if not builder.has_import("com.sun.star.uno.XWeak"):
        builder.auto_add_interface("com.sun.star.uno.XWeak")
    if not builder.has_import("com.sun.star.lang.XTypeProvider"):
        builder.auto_add_interface("com.sun.star.lang.XTypeProvider")
    if not builder.has_import("com.sun.star.lang.XServiceInfo"):
        builder.auto_add_interface("com.sun.star.lang.XServiceInfo")
    if not builder.has_import("ooodev.utils.partial.interface_partial.InterfacePartial"):
        builder.add_import(
            name="ooodev.utils.partial.interface_partial.InterfacePartial",
            optional=False,
            init_kind=InitKind.NONE,
            check_kind=CheckKind.NONE,
        )
    if not builder.has_import("ooodev.utils.partial.qi_partial.QiPartial"):
        builder.add_import(
            name="ooodev.utils.partial.qi_partial.QiPartial",
            uno_name="com.sun.star.uno.XInterface",
            optional=True,
            init_kind=InitKind.COMPONENT,
            check_kind=CheckKind.INTERFACE,
        )

    if not builder.has_import("ooodev.events.partial.events_partial.EventsPartial"):
        builder.insert_import(
            idx=0,
            name="ooodev.events.partial.events_partial.EventsPartial",
            optional=False,
            init_kind=InitKind.NONE,
            check_kind=CheckKind.NONE,
        )


def builder_add_service_defaults(builder: DefaultBuilder) -> None:
    """
    Adds the default service imports for the builder.

    Args:
        builder (DefaultBuilder): A Builder instance.

    Note:
        Not currently adding anything. May be used in the future.
    """
    pass


def builder_add_interface_defaults(builder: DefaultBuilder) -> None:
    """
    Adds the default interface imports for the builder.

    Args:
        builder (DefaultBuilder): A Builder instance.

    Note:
        Not currently adding anything. May be used in the future.
    """
    pass


def builder_add_property_change_implement(builder: DefaultBuilder) -> None:
    if not builder.has_import("ooodev.adapter.beans.property_change_implement.PropertyChangeImplement"):
        builder.add_import(
            name="ooodev.adapter.beans.property_change_implement.PropertyChangeImplement",
            uno_name="com.sun.star.beans.XPropertySet",
            optional=True,
            init_kind=InitKind.COMPONENT,
            check_kind=CheckKind.INTERFACE,
        )


def builder_add_property_veto_implement(builder: DefaultBuilder) -> None:
    if not builder.has_import("ooodev.adapter.beans.vetoable_change_implement.VetoableChangeImplement"):
        builder.add_import(
            name="ooodev.adapter.beans.vetoable_change_implement.VetoableChangeImplement",
            uno_name="com.sun.star.beans.XPropertySet",
            optional=True,
            init_kind=InitKind.COMPONENT,
            check_kind=CheckKind.INTERFACE,
        )


def builder_add_lo_inst_props_partial(builder: DefaultBuilder) -> None:
    if not builder.has_import("ooodev.utils.partial.lo_inst_props_partial.LoInstPropsPartial"):
        builder.add_import(
            name="ooodev.utils.partial.lo_inst_props_partial.LoInstPropsPartial",
            uno_name="",
            optional=False,
            init_kind=InitKind.LO_INST,
            check_kind=CheckKind.NONE,
        )


def builder_add_component_prop(builder: DefaultBuilder) -> None:
    if not builder.has_import("ooodev.adapter.component_prop.ComponentProp"):
        builder.add_import(
            name="ooodev.adapter.component_prop.ComponentProp",
            uno_name="",
            optional=False,
            init_kind=InitKind.COMPONENT,
            check_kind=CheckKind.NONE,
        )
