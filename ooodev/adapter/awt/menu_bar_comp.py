from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from com.sun.star.awt import XMenuBar
from ooodev.adapter._helper.builder import builder_helper
from ooodev.adapter._helper.builder.comp_defaults_partial import CompDefaultsPartial
from ooodev.adapter.component_prop import ComponentProp
from ooodev.utils.builder.default_builder import DefaultBuilder
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.awt.menu_bar_partial import MenuBarPartial
from ooodev.adapter.awt.menu_events import MenuEvents
from ooodev.adapter.lang.service_info_partial import ServiceInfoPartial
from ooodev.events.args.listener_event_args import ListenerEventArgs

if TYPE_CHECKING:
    from com.sun.star.awt import MenuBar
    from ooodev.loader.inst.lo_inst import LoInst


class _MenuBarComp(ComponentProp):

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
        return ("com.sun.star.awt.MenuBar",)


class MenuBarComp(_MenuBarComp, MenuBarPartial, ServiceInfoPartial, MenuEvents):
    """
    Class for managing MenuBar Component.
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
            name="ooodev.adapter.awt.menu_bar_comp.MenuBarComp",
            base_class=_MenuBarComp,
        )
        return inst

    def __init__(self, component: XMenuBar) -> None:
        """
        Constructor

        Args:
            component (UnoControlDialog): UNO Component that supports `com.sun.star.awt.MenuBar`` service.
        """
        # this it not actually called as __new__ is overridden
        pass

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.awt.MenuBar",)

    # endregion Overrides

    @classmethod
    def from_lo(cls, lo_inst: LoInst | None = None) -> MenuBarComp:
        """
        Creates a new instance from Lo.

        Args:
            lo_inst (LoInst, optional): LoInst, Defaults to ``Lo.current_lo``.

        Returns:
            MenuBarComp: The instance.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.loader import lo as mLo

        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        inst = lo_inst.create_instance_mcf(XMenuBar, "com.sun.star.awt.MenuBar", raise_err=True)  # type: ignore
        return cls(inst)

    # region Properties

    @property
    def component(self) -> MenuBar:
        """MenuBar Component"""
        # pylint: disable=no-member
        return cast("MenuBar", self._ComponentBase__get_component())  # type: ignore

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
    builder.auto_interface()
    builder.set_omit("ooodev.adapter.awt.menu_partial.MenuPartial")
    builder.add_event(
        module_name="ooodev.adapter.awt.menu_events",
        class_name="MenuEvents",
        uno_name="com.sun.star.awt.XMenuBar",
        optional=True,
    )
    return builder
