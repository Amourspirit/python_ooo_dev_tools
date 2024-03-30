from __future__ import annotations
from typing import Any, cast, List, TYPE_CHECKING, Tuple
import uno

from com.sun.star.awt import XUnoControlContainer
from com.sun.star.awt import XTabController
from ooodev.adapter.awt.tab_controller_comp import TabControllerComp
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class UnoControlContainerPartial:
    """
    Partial class for XUnoControlContainer.
    """

    def __init__(self, component: XUnoControlContainer, interface: UnoInterface | None = XUnoControlContainer) -> None:
        """
        Constructor

        Args:
            component (XUnoControlContainer): UNO Component that implements ``com.sun.star.awt.XUnoControlContainer`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XUnoControlContainer``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XUnoControlContainer
    def add_tab_controller(self, tab_controller: XTabController | TabControllerComp) -> None:
        """
        Adds a single tab controller.

        Args:
            tab_controller (XTabController | TabControllerComp): Tab controller to be added.
        """
        if mLo.Lo.is_uno_interfaces(tab_controller, XTabController):
            self.__component.addTabController(tab_controller)  # type: ignore
        else:
            self.__component.addTabController(tab_controller.component)  # type: ignore

    def get_tab_controllers(self) -> Tuple[TabControllerComp, ...]:
        """
        Gets all currently specified tab controllers.

        Returns:
            Tuple[TabControllerComp, ...]: Tab controllers
        """
        items = self.__component.getTabControllers()
        return tuple(TabControllerComp(item) for item in items) if items else ()

    def remove_tab_controller(self, tab_controller: XTabController | TabControllerComp) -> None:
        """
        removes a single tab controller.
        """
        if mLo.Lo.is_uno_interfaces(tab_controller, XTabController):
            self.__component.removeTabController(tab_controller)  # type: ignore
        else:
            self.__component.removeTabController(tab_controller.component)  # type: ignore

    def set_tab_controllers(self, *tab_controllers: XTabController | TabControllerComp) -> None:
        """
        Sets a set of tab controllers.

        Args:
            tab_controllers (XTabController | TabControllerComp): One or more Tab controllers to be set.
        """
        if not tab_controllers:
            return
        items = cast(
            List[XTabController],
            [item if mLo.Lo.is_uno_interfaces(item, XTabController) else item.component for item in tab_controllers],  # type: ignore
        )
        self.__component.setTabControllers(tuple(items))

    # endregion XUnoControlContainer
