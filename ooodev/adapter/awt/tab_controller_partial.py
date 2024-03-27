from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno

from com.sun.star.awt import XTabController

from ooodev.adapter.awt.control_container_comp import ControlContainerComp
from ooodev.adapter.awt.tab_controller_model_comp import TabControllerModelComp
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils import info as mInfo

if TYPE_CHECKING:
    from com.sun.star.awt import XControl
    from com.sun.star.awt import XTabControllerModel
    from com.sun.star.awt import XControlContainer
    from ooodev.utils.type_var import UnoInterface


class TabControllerPartial:
    """
    Partial class for XTabController.
    """

    def __init__(self, component: XTabController, interface: UnoInterface | None = XTabController) -> None:
        """
        Constructor

        Args:
            component (XTabController): UNO Component that implements ``com.sun.star.awt.XTabController`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XTabController``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XTabController
    def activate_first(self) -> None:
        """
        Sets the focus to the first control that can be reached with the TAB key.
        """
        self.__component.activateFirst()

    def activate_last(self) -> None:
        """
        Sets the focus to the last control that can be reached with the TAB key.
        """
        self.__component.activateLast()

    def activate_tab_order(self) -> None:
        """
        Activates tab order.
        """
        self.__component.activateTabOrder()

    def auto_tab_order(self) -> None:
        """
        Enables automatic tab order.
        """
        self.__component.autoTabOrder()

    def get_container(self) -> ControlContainerComp:
        """
        Gets the control container.

        Returns:
            ControlContainerComp: The control container.
        """
        return ControlContainerComp(self.__component.getContainer())

    def get_controls(self) -> Tuple[XControl, ...]:
        """
        Gets all controls of the control container.

        Returns:
            Tuple[XControl, ...]: The controls.
        """
        return self.__component.getControls()

    def get_model(self) -> TabControllerModelComp:
        """
        Gets the tab controller model.

        Returns:
            TabControllerModelComp: The model.
        """
        return TabControllerModelComp(self.__component.getModel())

    def set_container(self, container: XControlContainer | ControlContainerComp) -> None:
        """
        Set the control container.

        Args:
            container (XControlContainer | ControlContainerComp): The control container.
        """
        if mInfo.Info.is_instance(container, ControlContainerComp):
            self.__component.setContainer(container.component)
        else:
            self.__component.setContainer(container)  # type: ignore

    def set_model(self, model: XTabControllerModel | TabControllerModelComp) -> None:
        """
        Sets the tab controller model.

        Args:
            model (XTabControllerModel | TabControllerModelComp): The model.
        """
        if mInfo.Info.is_instance(model, TabControllerModelComp):
            self.__component.setModel(model.component)
        else:
            self.__component.setModel(model)  # type: ignore

    # endregion XTabController
