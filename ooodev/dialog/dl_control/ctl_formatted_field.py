# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno  # pylint: disable=unused-import

from ooodev.mock import mock_g
from ooodev.adapter.awt.uno_control_formatted_field_model_partial import UnoControlFormattedFieldModelPartial
from ooodev.adapter.awt.spin_events import SpinEvents
from ooodev.adapter.awt.text_events import TextEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind
from ooodev.dialog.dl_control.ctl_base import DialogControlBase, _create_control

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlFormattedField  # service
    from com.sun.star.awt import UnoControlFormattedFieldModel  # service
    from com.sun.star.awt import XWindowPeer
    from ooodev.dialog.dl_control.model.model_formatted_field import ModelFormattedField
    from ooodev.dialog.dl_control.view.view_formatted_field import ViewFormattedField
# endregion imports


class CtlFormattedField(DialogControlBase, UnoControlFormattedFieldModelPartial, SpinEvents, TextEvents):
    """Class for Formatted Field Control"""

    # pylint: disable=unused-argument

    # region init
    def __init__(self, ctl: UnoControlFormattedField) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlFormattedField): Button Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        DialogControlBase.__init__(self, ctl)
        UnoControlFormattedFieldModelPartial.__init__(self, component=self.get_model())
        generic_args = self._get_generic_args()
        # EventArgs.event_data will contain the ActionEvent
        SpinEvents.__init__(self, trigger_args=generic_args, cb=self._on_spin_events_listener_add_remove)
        TextEvents.__init__(self, trigger_args=generic_args, cb=self._on_text_events_listener_add_remove)
        self._model_ex = None
        self._view_ex = None

    # endregion init

    def __repr__(self) -> str:
        if hasattr(self, "name"):
            return f"CtlFormattedField({self.name})"
        return "CtlFormattedField"

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

    def get_view_ctl(self) -> UnoControlFormattedField:
        return cast("UnoControlFormattedField", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlFormattedField``"""
        return "com.sun.star.awt.UnoControlFormattedField"

    def get_model(self) -> UnoControlFormattedFieldModel:
        """Gets the Model for the control"""
        return cast("UnoControlFormattedFieldModel", self.get_view_ctl().getModel())

    def get_control_kind(self) -> DialogControlKind:
        """Gets the control kind. Returns ``DialogControlKind.FORMATTED_TEXT``"""
        return DialogControlKind.FORMATTED_TEXT

    def get_control_named_kind(self) -> DialogControlNamedKind:
        """Gets the control named kind. Returns ``DialogControlNamedKind.FORMATTED_TEXT``"""
        return DialogControlNamedKind.FORMATTED_TEXT

    # endregion Overrides

    # region Static Methods
    @staticmethod
    def create(win: XWindowPeer, **kwargs: Any) -> "CtlFormattedField":
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
            CtlFormattedField: New instance of the control.

        Note:
            The `UnoControlDialogElement <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1awt_1_1UnoControlDialogElement.html>`__
            interface is not included when creating the control with a window peer.
        """
        ctrl = _create_control("com.sun.star.awt.UnoControlFormattedFieldModel", win, **kwargs)
        return CtlFormattedField(ctl=ctrl)

    # endregion Static Methods

    # region Properties

    @property
    def model(self) -> UnoControlFormattedFieldModel:
        # pylint: disable=no-member
        return cast("UnoControlFormattedFieldModel", super().model)

    @property
    def model_ex(self) -> ModelFormattedField:
        """
        Gets the extended Model for the control.

        This is a wrapped instance for the model property.
        It add some additional properties and methods to the model.
        """
        # pylint: disable=no-member
        if self._model_ex is None:
            # pylint: disable=import-outside-toplevel
            # pylint: disable=redefined-outer-name
            from ooodev.dialog.dl_control.model.model_formatted_field import ModelFormattedField

            self._model_ex = ModelFormattedField(self.model)
        return self._model_ex

    @property
    def value(self) -> Any:
        """
        Gets/Sets the value.

        Same as ``effective_value`` property

        This may be a numeric value (float) or a string, depending on the formatting of the field.
        """
        return self.effective_value

    @value.setter
    def value(self, value: Any) -> None:
        self.effective_value = value

    @property
    def view(self) -> UnoControlFormattedField:
        # pylint: disable=no-member
        return cast("UnoControlFormattedField", super().view)

    @property
    def view_ex(self) -> ViewFormattedField:
        """
        Gets the extended View for the control.

        This is a wrapped instance for the view property.
        It add some additional properties and methods to the view.
        """
        # pylint: disable=no-member
        if self._view_ex is None:
            # pylint: disable=import-outside-toplevel
            # pylint: disable=redefined-outer-name
            from ooodev.dialog.dl_control.view.view_formatted_field import ViewFormattedField

            self._view_ex = ViewFormattedField(self.view)
        return self._view_ex

    # endregion Properties


if mock_g.FULL_IMPORT:
    from ooodev.dialog.dl_control.model.model_formatted_field import ModelFormattedField
    from ooodev.dialog.dl_control.view.view_formatted_field import ViewFormattedField
