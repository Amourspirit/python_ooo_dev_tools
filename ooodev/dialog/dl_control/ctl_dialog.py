# region imports
from __future__ import annotations
from typing import Any, cast, Tuple, TYPE_CHECKING
import uno  # pylint: disable=unused-import
from com.sun.star.awt import XControlContainer
from ooo.dyn.awt.pos_size import PosSizeEnum
from ooodev.adapter.awt.top_window_events import TopWindowEvents
from ooodev.adapter.awt.window_events import WindowEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import lo as mLo

from .ctl_base import CtlListenerBase

if TYPE_CHECKING:
    from com.sun.star.awt import XControl
    from com.sun.star.awt import UnoControlDialog  # service
    from com.sun.star.awt import UnoControlDialogModel  # service
    from com.sun.star.awt import Rectangle  # struct
# endregion imports


class CtlDialog(CtlListenerBase, TopWindowEvents, WindowEvents):
    """Class for Button Control"""

    # pylint: disable=unused-argument
    # region init
    def __init__(self, ctl: UnoControlDialog) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlDialog): Button Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        CtlListenerBase.__init__(self, ctl)
        generic_args = self._get_generic_args()
        # EventArgs.event_data will contain the ActionEvent
        TopWindowEvents.__init__(self, trigger_args=generic_args, cb=self._on_top_window_events_listener_add_remove)
        WindowEvents.__init__(self, trigger_args=generic_args, cb=self._on_window_events_listener_add_remove)

    # endregion init

    # region Lazy Listeners
    def _on_top_window_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        key = cast(str, event.source)
        if self._has_listener(key):
            return
        self.control.addTopWindowListener(self.events_listener_top_window)
        self._add_listener(key)

    def _on_window_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        key = cast(str, event.source)
        if self._has_listener(key):
            return
        self.control.addWindowListener(self.events_listener_window)
        self._add_listener(key)

    # endregion Lazy Listeners

    # region Overrides
    def get_view_ctl(self) -> UnoControlDialog:
        return cast("UnoControlDialog", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlDialog``"""
        return "com.sun.star.awt.UnoControlDialog"

    def get_model(self) -> UnoControlDialogModel:
        """Gets the Model for the control"""
        return cast("UnoControlDialogModel", self.get_view_ctl().getModel())

    # endregion Overrides

    # region Dialog Methods
    def execute(self) -> int:
        """
        Runs the dialog modally: shows it, and waits for the execution to end.

        Returns:
            int: Returns an exit code (e.g., indicating the button that was used to end the execution).
        """
        return self.control.execute()

    def dispose(self) -> None:
        """
        Disposes the dialog

        Returns:
            None:
        """
        self.control.dispose()

    def get_dialog_controls_arr(self) -> Tuple[XControl, ...]:
        """
        Gets all controls for the dialog

        Returns:
            Tuple[XControl, ...]: controls
        """
        ctrl_con = mLo.Lo.qi(XControlContainer, self.control, True)
        return ctrl_con.getControls()

    def find_control(self, name: str) -> XControl:
        """
        Finds control by name

        Args:
            name (str): Name to find

        Returns:
            XControl: Control
        """
        ctrl_con = mLo.Lo.qi(XControlContainer, self.control, True)
        return ctrl_con.getControl(name)

    def set_pos_size(self, x: int, y: int, width: int, height: int, flags: int | PosSizeEnum) -> None:
        """
        sets the outer bounds of the window.

        Args:
            x (int): X coordinate.
            y (int): Y coordinate.
            Width (int): _description_
            Height (int): _description_
            Flags (int | PosSizeEnum): _description_
        """
        self.control.setPosSize(x, y, width, height, int(flags))

    def get_pos_size(self) -> Rectangle:
        return self.control.getPosSize()

    def set_visible(self, visible: bool) -> None:
        """
        Shows or hides the window depending on the parameter.
        """
        self.control.setVisible(visible)

    # endregion Dialog Methods

    # region Properties
    @property
    def control(self) -> UnoControlDialog:
        """Returns the control"""
        return self.get_view_ctl()

    @property
    def model(self) -> UnoControlDialogModel:
        """Gets the Model for the control"""
        return self.get_model()

    @property
    def title(self) -> str:
        """Gets the title of the dialog"""
        return self.control.getTitle()

    @title.setter
    def title(self, value: str) -> None:
        """Sets the title of the dialog"""
        self.control.setTitle(value)

    # endregion Properties
