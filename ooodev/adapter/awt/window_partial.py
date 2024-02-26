from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno

from com.sun.star.awt import XWindow
from ooo.dyn.awt.pos_size import PosSize

from ooodev.utils.type_var import UnoInterface
from ooodev.adapter.awt.control_partial import ControlPartial

if TYPE_CHECKING:
    from com.sun.star.awt import XFocusListener
    from com.sun.star.awt import XKeyListener
    from com.sun.star.awt import XMouseListener
    from com.sun.star.awt import XMouseMotionListener
    from com.sun.star.awt import XPaintListener
    from com.sun.star.awt import XWindowListener
    from com.sun.star.awt import Rectangle  # Struct
    from ooodev.units.unit_obj import UnitT


class WindowPartial(ControlPartial):
    """
    Partial Class for XWindow.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XWindow, interface: UnoInterface | None = XWindow) -> None:
        """
        Constructor

        Args:
            component (XWindow): UNO Component that implements ``com.sun.star.awt.XWindow`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XWindow``.
        """
        ControlPartial.__init__(self, component=component, interface=interface)  # type: ignore
        self.__component = component

    # region XWindow
    def add_focus_listener(self, listener: XFocusListener) -> None:
        """
        Adds a focus listener to the object.
        """
        self.__component.addFocusListener(listener)

    def add_key_listener(self, listener: XKeyListener) -> None:
        """
        Adds a key listener to the object.
        """
        self.__component.addKeyListener(listener)

    def add_mouse_listener(self, listener: XMouseListener) -> None:
        """
        Adds a mouse listener to the object.
        """
        self.__component.addMouseListener(listener)

    def add_mouse_motion_listener(self, listener: XMouseMotionListener) -> None:
        """
        Adds a mouse motion listener to the object.
        """
        self.__component.addMouseMotionListener(listener)

    def add_paint_listener(self, listener: XPaintListener) -> None:
        """
        Adds a paint listener to the object.
        """
        self.__component.addPaintListener(listener)

    def add_window_listener(self, listener: XWindowListener) -> None:
        """
        Adds a window listener to the object.
        """
        self.__component.addWindowListener(listener)

    def get_pos_size(self) -> Rectangle:
        """
        Gets the outer bounds of the window.
        """
        return self.__component.getPosSize()

    def remove_focus_listener(self, listener: XFocusListener) -> None:
        """
        Removes the specified focus listener from the listener list.
        """
        self.__component.removeFocusListener(listener)

    def remove_key_listener(self, listener: XKeyListener) -> None:
        """
        Removes the specified key listener from the listener list.
        """
        self.__component.removeKeyListener(listener)

    def remove_mouse_listener(self, listener: XMouseListener) -> None:
        """
        Removes the specified mouse listener from the listener list.
        """
        self.__component.removeMouseListener(listener)

    def remove_mouse_motion_listener(self, listener: XMouseMotionListener) -> None:
        """
        Removes the specified mouse motion listener from the listener list.
        """
        self.__component.removeMouseMotionListener(listener)

    def remove_paint_listener(self, listener: XPaintListener) -> None:
        """
        Removes the specified paint listener from the listener list.
        """
        self.__component.removePaintListener(listener)

    def remove_window_listener(self, listener: XWindowListener) -> None:
        """
        Removes the specified window listener from the listener list.
        """
        self.__component.removeWindowListener(listener)

    def set_enable(self, enable: bool) -> None:
        """
        Enables or disables the window depending on the parameter.
        """
        self.__component.setEnable(enable)

    def set_focus(self) -> None:
        """
        sets the focus to the window.
        """
        self.__component.setFocus()

    def set_pos_size(
        self, x: int | UnitT, y: int | UnitT, width: int | UnitT, height: int | UnitT, flags: int = PosSize.POSSIZE
    ) -> None:
        """
        Sets the outer bounds of the window.

        Args:
            x (int, UnitT): The x-coordinate of the window. In ``1/100 mm`` or ``UnitT``.
            y (int, UnitT): The y-coordinate of the window. In ``1/100 mm`` or ``UnitT``.
            width (int, UnitT): The width of the window. In ``1/100 mm`` or ``UnitT``.
            height (int, UnitT): The height of the window. In ``1/100 mm`` or ``UnitT``.
            flags (int, UnitT): A combination of ``com.sun.star.awt.PosSize`` flags. Default set to ``PosSize.POSSIZE``.

        Returns:
            None:

        See Also:
            `com.sun.star.awt.PosSize <https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1awt_1_1PosSize.html>`__
        """
        try:
            x_arg = cast(int, x.get_value_mm100())  # type: ignore
        except AttributeError:
            x_arg = cast(int, x)
        try:
            y_arg = cast(int, y.get_value_mm100())  # type: ignore
        except AttributeError:
            y_arg = cast(int, y)
        try:
            width_arg = cast(int, width.get_value_mm100())  # type: ignore
        except AttributeError:
            width_arg = cast(int, width)
        try:
            height_arg = cast(int, height.get_value_mm100())  # type: ignore
        except AttributeError:
            height_arg = cast(int, height)

        self.__component.setPosSize(x_arg, y_arg, width_arg, height_arg, int(flags))

    def set_visible(self, visible: bool) -> None:
        """
        Shows or hides the window depending on the parameter.
        """
        self.__component.setVisible(visible)

    # endregion XWindow
