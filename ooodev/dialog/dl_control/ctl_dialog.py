# region imports
from __future__ import annotations
from typing import Any, cast, overload, Tuple, TYPE_CHECKING
import uno  # pylint: disable=unused-import
from com.sun.star.awt import XToolkit2
from com.sun.star.awt import XControlContainer
from com.sun.star.awt import XWindowPeer
from ooo.dyn.awt.pos_size import PosSize

from ooodev.units.unit_px import UnitPX
from ooodev.units.unit_app_font_height import UnitAppFontHeight
from ooodev.units.unit_app_font_width import UnitAppFontWidth
from ooodev.units.unit_app_font_x import UnitAppFontX
from ooodev.units.unit_app_font_y import UnitAppFontY
from ooodev.adapter.awt.top_window_events import TopWindowEvents
from ooodev.adapter.awt.window_events import WindowEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.loader import lo as mLo
from ooodev.adapter.awt.uno_control_dialog_comp import UnoControlDialogComp
from ooodev.adapter.awt.unit_conversion_partial import UnitConversionPartial

from ooodev.dialog.dl_control.ctl_base import CtlListenerBase
from ooodev.dialog.dl_control.model.model_dialog import ModelDialog

if TYPE_CHECKING:
    from com.sun.star.awt import XControl
    from com.sun.star.awt import UnoControlDialog  # service
    from com.sun.star.awt import UnoControlDialogModel  # service
    from com.sun.star.awt import XToolkit
    from ooodev.units.unit_obj import UnitT
# endregion imports


class CtlDialog(UnoControlDialogComp, CtlListenerBase, UnitConversionPartial, TopWindowEvents, WindowEvents):
    """Class for Dialog Control"""

    def __new__(cls, *args, **kwargs):
        return super().__new__(cls, *args, **kwargs)

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
        self._model_ex = None

    # endregion init

    def __repr__(self) -> str:
        if hasattr(self, "name"):
            return f"CtlDialog({self.name})"
        return "CtlDialog"

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

    def set_pos_size(
        self, x: int | UnitT, y: int | UnitT, width: int | UnitT, height: int | UnitT, flags: int = PosSize.POSSIZE
    ) -> None:
        """
        Sets the outer bounds of the window.

        Args:
            x (int, UnitT): The x-coordinate of the window. In ``Pixels`` or ``UnitT``.
            y (int, UnitT): The y-coordinate of the window. In ``Pixels`` or ``UnitT``.
            width (int, UnitT): The width of the window. In ``Pixels`` or ``UnitT``.
            height (int, UnitT): The height of the window. In ``Pixels`` or ``UnitT``.
            flags (int, UnitT): A combination of ``com.sun.star.awt.PosSize`` flags. Default set to ``PosSize.POSSIZE``.

        Returns:
            None:

        Note:
            The Model is in AppFont units where as the View is in Pixels.
            The values are converted from Pixels to AppFont units and assigned to the model.
            If the values are passed in as ``UnitAppFontX``, ``UnitAppFontY``, ``UnitAppFontWidth``, or ``UnitAppFontHeight``
            then they are assigned directly to the model without conversion.
        """

        def is_flag_set(value: int, flag: int) -> bool:
            return (value & flag) != 0

        if is_flag_set(flags, PosSize.X):
            if isinstance(x, UnitAppFontX):
                self.model.PositionX = int(x)  # type: ignore
            else:
                px_x = UnitPX.from_unit_val(x)
                self.model.PositionX = int(UnitAppFontX.from_unit_val(px_x))  # type: ignore

        if is_flag_set(flags, PosSize.Y):
            if isinstance(y, UnitAppFontY):
                self.model.PositionY = int(y)  # type: ignore
            else:
                px_y = UnitPX.from_unit_val(y)
                self.model.PositionY = int(UnitAppFontY.from_unit_val(px_y))  # type: ignore

        if is_flag_set(flags, PosSize.WIDTH):
            if isinstance(width, UnitAppFontWidth):
                self.model.Width = int(width)
            else:
                px_width = UnitPX.from_unit_val(width)
                self.model.Width = int(UnitAppFontWidth.from_unit_val(px_width))

        if is_flag_set(flags, PosSize.HEIGHT):
            if isinstance(height, UnitAppFontHeight):
                self.model.Height = int(height)
            else:
                px_height = UnitPX.from_unit_val(height)
                self.model.Height = int(UnitAppFontHeight.from_unit_val(px_height))

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
    def model_ex(self) -> ModelDialog:
        """
        Gets the extended Model for the control.

        This is a wrapped instance for the model property.
        It add some additional properties and methods to the model.
        """
        if self._model_ex is None:
            self._model_ex = ModelDialog(self.get_model())
        return self._model_ex

    @property
    def title(self) -> str:
        """Gets the title of the dialog"""
        return self.get_title()

    @title.setter
    def title(self, value: str) -> None:
        """Sets the title of the dialog"""
        self.set_title(value)

    # endregion Properties
