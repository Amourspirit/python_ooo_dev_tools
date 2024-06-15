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

    keys = {
        "com.sun.star.uno.XWeak",
        "com.sun.star.lang.XTypeProvider",
        "com.sun.star.lang.XServiceInfo",
    }

    for key in keys:
        if not builder.has_import(key) and not builder.has_omit(key):
            builder.auto_add_interface(key)

    key = "ooodev.utils.partial.interface_partial.InterfacePartial"
    if not builder.has_import(key) and not builder.has_omit(key):
        builder.add_import(
            name=key,
            optional=False,
            init_kind=InitKind.NONE,
            check_kind=CheckKind.NONE,
        )

    key = "ooodev.utils.partial.qi_partial.QiPartial"
    if not builder.has_import(key) and not builder.has_omit(key):
        # QiPartial will always be added as a component because it is XInterface related.
        # If this is set to optional it causes a issue with some classes reporting they have no qi attribute.
        builder.add_import(
            name=key,
            uno_name="com.sun.star.uno.XInterface",
            optional=False,
            init_kind=InitKind.COMPONENT,
            check_kind=CheckKind.NONE,
        )

    key = "ooodev.events.partial.events_partial.EventsPartial"
    if not builder.has_import(key) and not builder.has_omit(key):
        builder.insert_import(
            idx=0,
            name=key,
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
    key = "ooodev.adapter.beans.property_change_implement.PropertyChangeImplement"
    if not builder.has_import(key) and not builder.has_omit(key):
        builder.add_import(
            name=key,
            uno_name="com.sun.star.beans.XPropertySet",
            optional=True,
            init_kind=InitKind.COMPONENT,
            check_kind=CheckKind.INTERFACE,
        )


def builder_add_property_veto_implement(builder: DefaultBuilder) -> None:
    key = "ooodev.adapter.beans.vetoable_change_implement.VetoableChangeImplement"
    if not builder.has_import(key) and not builder.has_omit(key):
        builder.add_import(
            name=key,
            uno_name="com.sun.star.beans.XPropertySet",
            optional=True,
            init_kind=InitKind.COMPONENT,
            check_kind=CheckKind.INTERFACE,
        )


def builder_add_lo_inst_props_partial(builder: DefaultBuilder) -> None:
    key = "ooodev.utils.partial.lo_inst_props_partial.LoInstPropsPartial"
    if not builder.has_import(key) and not builder.has_omit(key):
        builder.add_import(
            name=key,
            uno_name="",
            optional=False,
            init_kind=InitKind.LO_INST,
            check_kind=CheckKind.NONE,
        )


def builder_add_component_prop(builder: DefaultBuilder) -> None:
    key = "ooodev.adapter.component_prop.ComponentProp"
    if not builder.has_import(key) and not builder.has_omit(key):
        builder.add_import(
            name=key,
            uno_name="",
            optional=False,
            init_kind=InitKind.COMPONENT,
            check_kind=CheckKind.NONE,
        )


def builder_add_the_dictionary_partial(builder: DefaultBuilder) -> None:
    key = "ooodev.utils.partial.the_dictionary_partial.TheDictionaryPartial"
    if not builder.has_import(key) and not builder.has_omit(key):
        builder.add_import(
            name=key,
            uno_name="",
            optional=False,
            init_kind=InitKind.NONE,
            check_kind=CheckKind.NONE,
        )
