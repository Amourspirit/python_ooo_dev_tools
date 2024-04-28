from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from ooodev.adapter.component_prop import ComponentProp
from com.sun.star.container import XEnumerationAccess
from com.sun.star.container import XContainer
from ooodev.adapter.awt.window_comp import WindowComp
from ooodev.adapter.container.container_partial import ContainerPartial
from ooodev.adapter.container.enumeration_access_partial import EnumerationAccessPartial
from ooodev.adapter.ui.action_trigger_container_comp import ActionTriggerContainerComp
from ooodev.adapter.view.selection_supplier_comp import SelectionSupplierComp
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.awt import Point  # struct
    from com.sun.star.ui import ContextMenuExecuteEvent  # struct
    from ooodev.utils.builder.default_builder import DefaultBuilder


class ContextMenuExecuteEventComp(ComponentProp):
    """
    Class for managing ContextMenuExecuteEvent Component.
    """

    # pylint: disable=unused-argument
    # def __new__(cls, component: Any, *args, **kwargs):
    #     builder = get_builder(component=component)
    #     builder_helper.builder_add_comp_defaults(builder)

    #     builder_only = kwargs.get("_builder_only", False)
    #     if builder_only:
    #         # cast to prevent type checker error
    #         return cast(Any, builder)
    #     inst = builder.build_class(
    #         name="ooodev.adapter.ui.context_menu_execute_event_comp.ContextMenuExecuteEventComp",
    #         base_class=_ContextMenuExecuteEventComp,
    #     )
    #     return inst

    def __init__(self, component: ContextMenuExecuteEvent) -> None:
        """
        Constructor

        Args:
            component (ContextMenuExecuteEvent): UNO ContextMenuExecuteEvent Component that supports ``com.sun.star.ui.ContextMenuExecuteEvent`` service.
        """
        ComponentProp.__init__(self, component)
        self.__action_trigger_container = None

    def __eq__(self, other: Any) -> bool:
        if not isinstance(other, ComponentProp):
            return False
        if self is other:
            return True
        if self.component is other.component:
            return True
        return self.component == other.component

    def __get_action_trigger_container(self) -> ActionTriggerContainerComp:
        """
        ActionTriggerContainer may also support XEnumerationAccess and XContainer interfaces.
        This method dynamically creates a new class that combines all the partial classes and the base class.
        """
        oth_bases = []

        def generate_class() -> type:
            nonlocal oth_bases
            bases = [ActionTriggerContainerComp] + oth_bases
            if len(bases) == 1:
                return bases[0]
            else:
                return type("ActionTriggerContainer", tuple(bases), {})

        comp = self.component.ActionTriggerContainer
        if mLo.Lo.is_uno_interfaces(comp, XEnumerationAccess):
            oth_bases.append(EnumerationAccessPartial)
        if mLo.Lo.is_uno_interfaces(comp, XContainer):
            oth_bases.append(ContainerPartial)
        clazz = generate_class()
        instance = clazz(comp)
        for t in oth_bases:
            t.__init__(instance, comp, None)
        return instance

    # region Properties
    @property
    def component(self) -> ContextMenuExecuteEvent:
        """ContextMenuExecuteEvent Component"""
        # pylint: disable=no-member
        return cast("ContextMenuExecuteEvent", self._ComponentBase__get_component())  # type: ignore

    @property
    def action_trigger_container(self) -> ActionTriggerContainerComp:
        """ActionTriggerContainer Component"""
        # need to dynamically create container
        if self.__action_trigger_container is None:
            self.__action_trigger_container = self.__get_action_trigger_container()
        return self.__action_trigger_container

    @property
    def execution_position(self) -> Point:
        """XTextRange Component"""
        return self.component.ExecutePosition

    @property
    def source_window(self) -> WindowComp:
        """Window Component"""
        return WindowComp(self.component.SourceWindow)

    @property
    def selection(self) -> SelectionSupplierComp:
        """SelectionSupplier Component"""
        return SelectionSupplierComp(self.component.Selection)

    @property
    def __class__(self):
        # pretend to be a ContextMenuExecuteEventComp class
        return ContextMenuExecuteEventComp

    # endregion Properties


def get_builder(component: Any) -> DefaultBuilder:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """
    # pylint: disable=import-outside-toplevel
    # pylint: disable=redefined-outer-name
    from ooodev.utils.builder.default_builder import DefaultBuilder

    builder = DefaultBuilder(component)

    builder.add_import(
        "ooodev.adapter.ui.context_menu_execute_event_comp.ContextMenuExecuteEventComp",
        optional=False,
        init_kind=1,
        check_kind=0,
    )
    return builder
