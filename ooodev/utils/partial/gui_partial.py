from __future__ import annotations
from typing import Any
import uno
from com.sun.star.frame import XController
from com.sun.star.frame import XModel
from com.sun.star.frame import XFrame
from com.sun.star.view import XControlAccess
from com.sun.star.view import XSelectionSupplier
from com.sun.star.frame import XDispatchProviderInterception

from ooodev.utils.inst.lo.lo_inst import LoInst
from ooodev.utils import gui as mGui


class GuiPartial:
    def __init__(self, component: Any, lo_inst: LoInst):
        self.__lo_inst = lo_inst  # may be used in future
        self.__component = component

    def get_current_controller(self) -> XController:
        """
        Gets controller from document.

        Returns:
            XController: controller.
        """
        model = self.__lo_inst.qi(XModel, self.__component, True)
        return model.getCurrentController()

    def get_frame(self) -> XFrame:
        """
        Gets frame from doc.

        Returns:
            XFrame: document frame.
        """
        controller = self.get_current_controller()
        return controller.getFrame()

    def get_control_access(self) -> XControlAccess:
        """
        Get control access from office document.

        Returns:
            XControlAccess: control access.
        """
        return self.__lo_inst.qi(XControlAccess, self.get_current_controller(), True)

    def get_selection_supplier(self) -> XSelectionSupplier:
        """
        Gets selection supplier

        Returns:
            XSelectionSupplier: Selection supplier
        """
        controller = self.get_current_controller()
        return self.__lo_inst.qi(XSelectionSupplier, controller, True)

    def get_dpi(self) -> XDispatchProviderInterception:
        """
        Gets Dispatch provider interception.

        Returns:
            XDispatchProviderInterception: Dispatch provider interception.
        """
        frame = self.get_frame()
        return self.__lo_inst.qi(XDispatchProviderInterception, frame, True)

    def activate(self) -> None:
        """
        Activates document window.
        """
        mGui.GUI.activate(self.__component)

    def maximize(self) -> None:
        """
        Maximizes document window.
        """
        mGui.GUI.maximize(self.__component)

    def minimize(self) -> None:
        """
        Minimizes document window.
        """
        mGui.GUI.minimize(self.__component)

    def add_item_to_toolbar(self, toolbar_name: str, item_name: str, im_fnm: str) -> None:
        """
        Add a user-defined icon and command to the start of the specified toolbar.

        Args:
            toolbar_name (str): toolbar name.
            item_name (str): item name.
            im_fnm (str): image file path.
        """
        mGui.GUI.add_item_to_toolbar(self.__component, toolbar_name, item_name, im_fnm)

    def show_menu_bar(self) -> None:
        """
        Shows the main menu bar.
        """
        mGui.GUI.show_menu_bar(self.__component)

    def hide_menu_bar(self) -> None:
        """
        Hides the main menu bar.
        """
        mGui.GUI.hide_menu_bar(self.__component)

    def hide_all_menu_bars(self) -> None:
        """
        Make all the toolbars invisible.
        """
        mGui.GUI.show_none(self.__component)

    def toggle_menu_bar(self) -> None:
        """
        Toggles the main menu bar.
        """
        mGui.GUI.toggle_menu_bar()
