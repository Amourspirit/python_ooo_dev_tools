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

        - ``InterfacePartial``
        - ``QiPartial``
    """
    builder.add_import(
        name="ooodev.utils.partial.interface_partial.InterfacePartial",
        optional=False,
        init_kind=InitKind.NONE,
        check_kind=CheckKind.NONE,
    )
    builder.add_import(
        name="ooodev.utils.partial.qi_partial.QiPartial",
        uno_name="com.sun.star.uno.XInterface",
        optional=True,
        init_kind=InitKind.COMPONENT_INTERFACE,
        check_kind=CheckKind.INTERFACE,
    )

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
        Adds the following imports:

        - ``ServiceInfoPartial``
    """
    builder.auto_add_interface("com.sun.star.lang.XServiceInfo")


def builder_add_interface_defaults(builder: DefaultBuilder) -> None:
    """
    Adds the default interface imports for the builder.

    Args:
        builder (DefaultBuilder): A Builder instance.

    Note:
        Not currently adding anything. May be used in the future.
    """
    pass
