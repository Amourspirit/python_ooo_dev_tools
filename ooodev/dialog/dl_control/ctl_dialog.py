# region imports
from __future__ import annotations
from typing import Any, cast, overload, Tuple, TYPE_CHECKING
import uno  # pylint: disable=unused-import
from com.sun.star.awt import XToolkit2
from com.sun.star.awt import XControlContainer
from com.sun.star.awt import XWindowPeer

from ooodev.adapter.awt.top_window_events import TopWindowEvents
from ooodev.adapter.awt.window_events import WindowEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.loader import lo as mLo
from ooodev.adapter.awt.index_container_comp import UnoControlDialogComp
from ooodev.adapter.awt.unit_conversion_partial import UnitConversionPartial

from ooodev.dialog.dl_control.ctl_base import CtlListenerBase

if TYPE_CHECKING:
    from com.sun.star.awt import XControl
    from com.sun.star.awt import UnoControlDialog  # service
    from com.sun.star.awt import UnoControlDialogModel  # service
    from com.sun.star.awt import XToolkit
# endregion imports


class CtlDialog(UnoControlDialogComp, CtlListenerBase, UnitConversionPartial, TopWindowEvents, WindowEvents):
    """Class for Dialog Control"""

    # pylint: disable=unused-argument
    # region init
    def __init__(self, ctl: UnoControlDialog) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlDialog): Dialog Control.
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        UnoControlDialogComp.__init__(self, component=ctl)
        CtlListenerBase.__init__(self, ctl)
        UnitConversionPartial.__init__(self, component=ctl)  # type: ignore
        generic_args = self._get_generic_args()
        # EventArgs.event_data will contain the ActionEvent
        TopWindowEvents.__init__(self, trigger_args=generic_args, cb=self._on_top_window_events_listener_add_remove)
        WindowEvents.__init__(self, trigger_args=generic_args, cb=self._on_window_events_listener_add_remove)

    # endregion init

    # region Lazy Listeners
    def _on_top_window_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.control.addTopWindowListener(self.events_listener_top_window)
        event.remove_callback = True

    def _on_window_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.control.addWindowListener(self.events_listener_window)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides
    def get_view(self) -> UnoControlDialog:
        view = UnoControlDialogComp.get_view(self)
        return cast("UnoControlDialog", view)

    def get_model(self) -> UnoControlDialogModel:
        """Gets the Model for the control"""
        model = UnoControlDialogComp.get_model(self)
        return cast("UnoControlDialogModel", model)

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlDialog``"""
        return "com.sun.star.awt.UnoControlDialog"

    # region create_peer()
    @overload
    def create_peer(self) -> None:
        """
        Creates a child window on the screen using the desktop window as the parent.
        """
        ...

    @overload
    def create_peer(self, toolkit: XToolkit) -> None:
        """
        Creates a child window on the screen using the desktop window as the parent.

        Args:
            toolkit (XToolkit): The toolkit to use.
        """
        ...

    @overload
    def create_peer(self, *, parent: XWindowPeer) -> None:
        """
        Creates a child window on the screen.

        Args:
            parent (XWindowPeer): The parent window.
        """
        ...

    @overload
    def create_peer(self, toolkit: XToolkit, parent: XWindowPeer) -> None:
        """
        Creates a child window on the screen.

        Args:
            toolkit (XToolkit): The toolkit to use.
            parent (XWindowPeer): The parent window.
        """
        ...

    def create_peer(self, toolkit: XToolkit | None = None, parent: XWindowPeer | None = None) -> None:
        """
        Creates a child window on the screen.

        If the parent is ``None``, then the desktop window of the toolkit is the parent.

        Args:
            toolkit (XToolkit, optional): The toolkit to use. Defaults to ``None``.
            parent (XWindowPeer, optional): The parent window. Defaults to ``None``.
        """
        if toolkit is None:
            # window = self.lo_inst.get_frame().getContainerWindow()
            toolkit = cast(XToolkit2, self.get_model().createInstance("com.sun.star.awt.Toolkit"))
            if toolkit is None:
                toolkit = self.lo_inst.create_instance_mcf(XToolkit2, "com.sun.star.awt.Toolkit", raise_err=True)
        if parent is None:
            parent = toolkit.getDesktopWindow()  # may still be None but that's ok
            if parent is None:
                parent = self.lo_inst.qi(XWindowPeer, self.lo_inst.get_frame().getContainerWindow(), True)
        UnoControlDialogComp.create_peer(self, toolkit, parent)  # type: ignore

    # endregion create_peer()

    # endregion Overrides

    # region Dialog Methods
    # def execute(self) -> int:
    #     """
    #     Runs the dialog modally: shows it, and waits for the execution to end.

    #     Returns:
    #         int: Returns an exit code (e.g., indicating the button that was used to end the execution).
    #     """
    #     return self.control.execute()

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

    # endregion Dialog Methods

    # region Properties
    @property
    def control(self) -> UnoControlDialog:
        """Returns the control"""
        return self.get_view()

    @property
    def model(self) -> UnoControlDialogModel:
        """Gets the Model for the control"""
        return self.get_model()

    @property
    def title(self) -> str:
        """Gets the title of the dialog"""
        return self.get_title()

    @title.setter
    def title(self, value: str) -> None:
        """Sets the title of the dialog"""
        self.set_title(value)

    # endregion Properties
