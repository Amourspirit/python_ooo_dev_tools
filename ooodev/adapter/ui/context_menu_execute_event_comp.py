from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from com.sun.star.container import XEnumerationAccess
from com.sun.star.container import XContainer
from ooodev.adapter.awt.window_comp import WindowComp
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.container.container_partial import ContainerPartial
from ooodev.adapter.container.enumeration_access_partial import EnumerationAccessPartial
from ooodev.adapter.ui.action_trigger_container_comp import ActionTriggerContainerComp
from ooodev.adapter.view.selection_supplier_comp import SelectionSupplierComp
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.awt import Point  # struct
    from com.sun.star.ui import ContextMenuExecuteEvent  # struct


class ContextMenuExecuteEventComp(ComponentBase):
    """
    Class for managing ContextMenuExecuteEvent Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: ContextMenuExecuteEvent) -> None:
        """
        Constructor

        Args:
            component (ContextMenuExecuteEvent): UNO ContextMenuExecuteEvent Component that supports ``com.sun.star.ui.ContextMenuExecuteEvent`` service.
        """
        # component is a struct
        ComponentBase.__init__(self, component)
        self.__action_trigger_container = None

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        # validated by mTextRangePartial.TextRangePartial
        return ()  # ("com.sun.star.ui.ContextMenuExecuteEvent",)

    # endregion Overrides

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

    # endregion Properties
