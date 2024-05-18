# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno  # pylint: disable=unused-import

from ooodev.mock import mock_g
from ooodev.adapter.awt.uno_control_scroll_bar_model_partial import UnoControlScrollBarModelPartial
from ooodev.adapter.awt.adjustment_events import AdjustmentEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind
from ooodev.dialog.dl_control.ctl_base import DialogControlBase, _create_control

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlScrollBar  # service
    from com.sun.star.awt import UnoControlScrollBarModel  # service
    from com.sun.star.awt import XWindowPeer
    from ooodev.dialog.dl_control.model.model_scroll_bar import ModelScrollBar
    from ooodev.dialog.dl_control.view.view_scroll_bar import ViewScrollBar

# endregion imports


class CtlScrollBar(DialogControlBase, UnoControlScrollBarModelPartial, AdjustmentEvents):
    """Class for Scroll Bar Control"""

    # pylint: disable=unused-argument

    # region init
    def __init__(self, ctl: UnoControlScrollBar) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlScrollBar): Scroll Bar Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        DialogControlBase.__init__(self, ctl)
        UnoControlScrollBarModelPartial.__init__(self, component=self.get_model())
        generic_args = self._get_generic_args()
        # EventArgs.event_data will contain the ActionEvent
        AdjustmentEvents.__init__(self, trigger_args=generic_args, cb=self._on_adjustment_events_listener_add_remove)
        self._model_ex = None
        self._view_ex = None

    # endregion init

    def __repr__(self) -> str:
        if hasattr(self, "name"):
            return f"CtlScrollBar({self.name})"
        return "CtlScrollBar"

    # region Lazy Listeners
    def _on_adjustment_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addAdjustmentListener(self.events_listener_adjustment)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides
    def get_view_ctl(self) -> UnoControlScrollBar:
        return cast("UnoControlScrollBar", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlScrollBar``"""
        return "com.sun.star.awt.UnoControlScrollBar"

    def get_model(self) -> UnoControlScrollBarModel:
        """Gets the Model for the control"""
        return cast("UnoControlScrollBarModel", self.get_view_ctl().getModel())

    def get_control_kind(self) -> DialogControlKind:
        """Gets the control kind. Returns ``DialogControlKind.SCROLL_BAR``"""
        return DialogControlKind.SCROLL_BAR

    def get_control_named_kind(self) -> DialogControlNamedKind:
        """Gets the control named kind. Returns ``DialogControlNamedKind.SCROLL_BAR``"""
        return DialogControlNamedKind.SCROLL_BAR

    # endregion Overrides

    # region Static Methods
    @staticmethod
    def create(win: XWindowPeer, **kwargs: Any) -> "CtlScrollBar":
        """
        Creates a new instance of the control.

        Keyword arguments are optional.
        Extra Keyword args are passed to the control as property values.

        Args:
            win (XWindowPeer): Parent Window

        Keyword Args:
            x (int, UnitT, optional): X Position in Pixels or UnitT.
            y (int, UnitT, optional): Y Position in Pixels or UnitT.
            width (int, UnitT, optional): Width in Pixels or UnitT.
            height (int, UnitT, optional): Height in Pixels or UnitT.

        Returns:
            CtlScrollBar: New instance of the control.

        Note:
            The `UnoControlDialogElement <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1awt_1_1UnoControlDialogElement.html>`__
            interface is not included when creating the control with a window peer.
        """
        ctrl = _create_control("com.sun.star.awt.UnoControlScrollBarModel", win, **kwargs)
        return CtlScrollBar(ctl=ctrl)

    # endregion Static Methods

    # region Properties

    @property
    def max_value(self) -> int:
        """Gets the maximum value of the scroll bar. Same as ``scroll_value_max`` property."""
        return self.scroll_value_max

    @max_value.setter
    def max_value(self, value: int) -> None:
        self.scroll_value_max = value

    @property
    def min_value(self) -> int:
        """Gets the minimum value of the scroll bar. Same as ``scroll_value_min`` property."""
        value = self.scroll_value_min
        return 0 if value is None else value

    @min_value.setter
    def min_value(self, value: int) -> None:
        self.scroll_value_min = value

    @property
    def model(self) -> UnoControlScrollBarModel:
        # pylint: disable=no-member
        return cast("UnoControlScrollBarModel", super().model)

    @property
    def model_ex(self) -> ModelScrollBar:
        """
        Gets the extended Model for the control.

        This is a wrapped instance for the model property.
        It add some additional properties and methods to the model.
        """
        # pylint: disable=no-member
        if self._model_ex is None:
            # pylint: disable=import-outside-toplevel
            # pylint: disable=redefined-outer-name
            from ooodev.dialog.dl_control.model.model_scroll_bar import ModelScrollBar

            self._model_ex = ModelScrollBar(self.model)
        return self._model_ex

    @property
    def value(self) -> int:
        """Gets or sets the current value of the scroll bar"""
        return self.view.getValue()

    @value.setter
    def value(self, value: int) -> None:
        self.view.setValue(value)

    @property
    def view(self) -> UnoControlScrollBar:
        # pylint: disable=no-member
        return cast("UnoControlScrollBar", super().view)

    @property
    def view_ex(self) -> ViewScrollBar:
        """
        Gets the extended View for the control.

        This is a wrapped instance for the view property.
        It add some additional properties and methods to the view.
        """
        # pylint: disable=no-member
        if self._view_ex is None:
            # pylint: disable=import-outside-toplevel
            # pylint: disable=redefined-outer-name
            from ooodev.dialog.dl_control.view.view_scroll_bar import ViewScrollBar

            self._view_ex = ViewScrollBar(self.view)
        return self._view_ex

    # endregion Properties


if mock_g.FULL_IMPORT:
    from ooodev.dialog.dl_control.model.model_scroll_bar import ModelScrollBar
    from ooodev.dialog.dl_control.view.view_scroll_bar import ViewScrollBar
