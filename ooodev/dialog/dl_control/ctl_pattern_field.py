# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import contextlib
import uno  # pylint: disable=unused-import

from ooodev.mock import mock_g
from ooodev.adapter.awt.uno_control_pattern_field_model_partial import UnoControlPatternFieldModelPartial
from ooodev.adapter.awt.spin_events import SpinEvents
from ooodev.adapter.awt.text_events import TextEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind
from ooodev.dialog.dl_control.ctl_base import DialogControlBase, _create_control

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlPatternField  # service
    from com.sun.star.awt import UnoControlPatternFieldModel  # service
    from com.sun.star.awt import XWindowPeer
    from ooodev.dialog.dl_control.model.model_pattern_field import ModelPatternField
    from ooodev.dialog.dl_control.view.view_pattern_field import ViewPatternField
# endregion imports


class CtlPatternField(DialogControlBase, UnoControlPatternFieldModelPartial, SpinEvents, TextEvents):
    """Class for Pattern Field Control"""

    # pylint: disable=unused-argument

    # region init
    def __init__(self, ctl: UnoControlPatternField) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlPatternField): Button Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        DialogControlBase.__init__(self, ctl)
        UnoControlPatternFieldModelPartial.__init__(self, component=self.get_model())
        generic_args = self._get_generic_args()
        # EventArgs.event_data will contain the ActionEvent
        SpinEvents.__init__(self, trigger_args=generic_args, cb=self._on_spin_events_listener_add_remove)
        TextEvents.__init__(self, trigger_args=generic_args, cb=self._on_text_events_listener_add_remove)
        self._model_ex = None
        self._view_ex = None

    # endregion init

    def __repr__(self) -> str:
        if hasattr(self, "name"):
            return f"CtlPatternField({self.name})"
        return "CtlPatternField"

    # region Lazy Listeners
    def _on_spin_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addSpinListener(self.events_listener_spin)
        event.remove_callback = True

    def _on_text_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addTextListener(self.events_listener_text)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides

    def get_view_ctl(self) -> UnoControlPatternField:
        return cast("UnoControlPatternField", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlPatternField``"""
        return "com.sun.star.awt.UnoControlPatternField"

    def get_model(self) -> UnoControlPatternFieldModel:
        """Gets the Model for the control"""
        return cast("UnoControlPatternFieldModel", self.get_view_ctl().getModel())

    def get_control_kind(self) -> DialogControlKind:
        """Gets the control kind. Returns ``DialogControlKind.PATTERN``"""
        return DialogControlKind.PATTERN

    def get_control_named_kind(self) -> DialogControlNamedKind:
        """Gets the control named kind. Returns ``DialogControlNamedKind.PATTERN``"""
        return DialogControlNamedKind.PATTERN

    # endregion Overrides

    # region Static Methods
    @staticmethod
    def create(win: XWindowPeer, **kwargs: Any) -> "CtlPatternField":
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
            CtlPatternField: New instance of the control.

        Note:
            The `UnoControlDialogElement <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1awt_1_1UnoControlDialogElement.html>`__
            interface is not included when creating the control with a window peer.
        """
        ctrl = _create_control("com.sun.star.awt.UnoControlPatternFieldModel", win, **kwargs)
        return CtlPatternField(ctl=ctrl)

    # endregion Static Methods

    # region Properties
    @property
    def model(self) -> UnoControlPatternFieldModel:
        # pylint: disable=no-member
        return cast("UnoControlPatternFieldModel", super().model)

    @property
    def model_ex(self) -> ModelPatternField:
        """
        Gets the extended Model for the control.

        This is a wrapped instance for the model property.
        It add some additional properties and methods to the model.
        """
        # pylint: disable=no-member
        if self._model_ex is None:
            # pylint: disable=import-outside-toplevel
            # pylint: disable=redefined-outer-name
            from ooodev.dialog.dl_control.model.model_pattern_field import ModelPatternField

            self._model_ex = ModelPatternField(self.model)
        return self._model_ex

    @property
    def view(self) -> UnoControlPatternField:
        # pylint: disable=no-member
        return cast("UnoControlPatternField", super().view)

    @property
    def view_ex(self) -> ViewPatternField:
        """
        Gets the extended View for the control.

        This is a wrapped instance for the view property.
        It add some additional properties and methods to the view.
        """
        # pylint: disable=no-member
        if self._view_ex is None:
            # pylint: disable=import-outside-toplevel
            # pylint: disable=redefined-outer-name
            from ooodev.dialog.dl_control.view.view_pattern_field import ViewPatternField

            self._view_ex = ViewPatternField(self.view)
        return self._view_ex

    # endregion Properties


if mock_g.FULL_IMPORT:
    from ooodev.dialog.dl_control.model.model_pattern_field import ModelPatternField
    from ooodev.dialog.dl_control.view.view_pattern_field import ViewPatternField
