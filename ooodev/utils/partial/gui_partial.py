from __future__ import annotations
from typing import Any, TYPE_CHECKING
import contextlib
import uno
from com.sun.star.awt import XTopWindow2
from com.sun.star.awt import XWindow2
from com.sun.star.frame import XController
from com.sun.star.frame import XDispatchProviderInterception
from com.sun.star.frame import XFrame
from com.sun.star.frame import XModel
from com.sun.star.view import XControlAccess
from com.sun.star.view import XSelectionSupplier

from ooodev.gui import gui as mGui
from ooodev.gui.comp.frame import Frame
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.context.lo_context import LoContext
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial


if TYPE_CHECKING:
    from com.sun.star.awt import XTopWindow


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

    def get_frame_comp(self) -> Frame:
        """
        Gets frame from doc as a FrameComp.

        Returns:
            FrameComp: document frame.
        """
        frm = self.get_frame()
        if frm is None:
            return None  # type: ignore
        if isinstance(self, LoInstPropsPartial):
            return Frame(frm, lo_inst=self.lo_inst)
        return Frame(frm)

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

    def get_selection(self) -> Any:
        """
        Gets selection.

        Returns:
            Any: Returns current selection or None.
        """
        with contextlib.suppress(Exception):
            return self.get_selection_supplier().getSelection()
        return None

    def get_dpi(self) -> XDispatchProviderInterception:
        """
        Gets Dispatch provider interception.

        Returns:
            XDispatchProviderInterception: Dispatch provider interception.
        """
        frame = self.get_frame()
        return self.__lo_inst.qi(XDispatchProviderInterception, frame, True)

    def get_top_window(self) -> XTopWindow | None:
        """
        Gets top window.

        Returns:
            XTopWindow | None: Top window or None if there is no Active Top Window.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.adapter.awt.toolkit_comp import ToolkitComp

        tk = ToolkitComp.from_lo(self.__lo_inst)
        return tk.get_active_top_window()

    def activate(self) -> None:
        """
        Activates document window.
        """

        def gui_activate():
            with LoContext(self.__lo_inst):
                mGui.GUI.activate(self.__component)

        frame = self.get_frame()
        if frame is not None:
            container = frame.getContainerWindow()
            if container is None:
                gui_activate()
                return
            x_win2 = self.__lo_inst.qi(XWindow2, container)
            if x_win2 is None:
                gui_activate()
                return
            top2 = self.__lo_inst.qi(XTopWindow2, container)
            if top2 is None:
                gui_activate()
                return
            if not x_win2.isVisible():
                x_win2.setVisible(True)
            if top2.IsMinimized:
                top2.IsMinimized = False
            x_win2.setFocus()
            top2.toFront()
        else:
            gui_activate()

    def maximize(self) -> None:
        """
        Maximizes document window.
        """
        with LoContext(self.__lo_inst):
            mGui.GUI.maximize(self.__component)

    def minimize(self) -> None:
        """
        Minimizes document window.
        """
        with LoContext(self.__lo_inst):
            mGui.GUI.minimize(self.__component)

    def add_item_to_toolbar(self, toolbar_name: str, item_name: str, im_fnm: str) -> None:
        """
        Add a user-defined icon and command to the start of the specified toolbar.

        Args:
            toolbar_name (str): toolbar name.
            item_name (str): item name.
            im_fnm (str): image file path.
        """
        with LoContext(self.__lo_inst):
            mGui.GUI.add_item_to_toolbar(self.__component, toolbar_name, item_name, im_fnm)

    def show_menu_bar(self) -> None:
        """
        Shows the main menu bar.
        """
        with LoContext(self.__lo_inst):
            mGui.GUI.show_menu_bar(self.__component)

    def hide_menu_bar(self) -> None:
        """
        Hides the main menu bar.
        """
        with LoContext(self.__lo_inst):
            mGui.GUI.hide_menu_bar(self.__component)

    def hide_all_menu_bars(self) -> None:
        """
        Make all the toolbars invisible.
        """
        with LoContext(self.__lo_inst):
            mGui.GUI.show_none(self.__component)

    def toggle_menu_bar(self) -> None:
        """
        Toggles the main menu bar.
        """
        with LoContext(self.__lo_inst):
            mGui.GUI.toggle_menu_bar()
