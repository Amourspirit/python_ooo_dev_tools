# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno  # pylint: disable=unused-import

# pylint: disable=useless-import-alias
from ooodev.mock import mock_g
from ooodev.adapter.awt.item_events import ItemEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind
from ooodev.adapter.awt.uno_control_check_box_model_partial import UnoControlCheckBoxModelPartial
from ooodev.dialog.dl_control.ctl_base import DialogControlBase, _create_control


if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlCheckBox  # service
    from com.sun.star.awt import UnoControlCheckBoxModel  # service
    from com.sun.star.awt import XWindowPeer
    from ooodev.dialog.dl_control.model.model_checkbox import ModelCheckbox
    from ooodev.dialog.dl_control.view.view_check_box import ViewCheckBox
# endregion imports


class CtlCheckBox(DialogControlBase, UnoControlCheckBoxModelPartial, ItemEvents):
    """Class for CheckBox Control"""

    # pylint: disable=unused-argument

    # region init
    def __init__(self, ctl: UnoControlCheckBox) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlCheckBox): Check Box Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        DialogControlBase.__init__(self, ctl)
        UnoControlCheckBoxModelPartial.__init__(self, self.model)
        generic_args = self._get_generic_args()
        # EventArgs.event_data will contain the ActionEvent
        ItemEvents.__init__(self, trigger_args=generic_args, cb=self._on_item_event_listener_add_remove)
        self._model_ex = None
        self._view_ex = None

    # endregion init

    def __repr__(self) -> str:
        if hasattr(self, "name"):
            return f"CtlCheckBox({self.name})"
        return "CtlCheckBox"

    # region Lazy Listeners
    def _on_item_event_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addItemListener(self.events_listener_item)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides
    def get_view_ctl(self) -> UnoControlCheckBox:
        return cast("UnoControlCheckBox", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlCheckBox``"""
        return "com.sun.star.awt.UnoControlCheckBox"

    def get_model(self) -> UnoControlCheckBoxModel:
        """Gets the Model for the control"""
        return cast("UnoControlCheckBoxModel", self.get_view_ctl().getModel())

    def get_control_kind(self) -> DialogControlKind:
        """Gets the control kind. Returns ``DialogControlKind.CHECKBOX``"""
        return DialogControlKind.CHECKBOX

    def get_control_named_kind(self) -> DialogControlNamedKind:
        """Gets the control named kind. Returns ``DialogControlNamedKind.CHECKBOX``"""
        return DialogControlNamedKind.CHECKBOX

    # endregion Overrides

    # region Static Methods
    @staticmethod
    def create(win: XWindowPeer, **kwargs: Any) -> "CtlCheckBox":
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
            CtlCheckBox: New instance of the control.

        Note:
            The `UnoControlDialogElement <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1awt_1_1UnoControlDialogElement.html>`__
            interface is not included when creating the control with a window peer.
        """
        ctrl = _create_control("com.sun.star.awt.UnoControlCheckBoxModel", win, **kwargs)
        return CtlCheckBox(ctl=ctrl)

    # endregion Static Methods

    # region Properties
    @property
    def model(self) -> UnoControlCheckBoxModel:
        # pylint: disable=no-member
        return cast("UnoControlCheckBoxModel", super().model)

    @property
    def model_ex(self) -> ModelCheckbox:
        """
        Gets the extended Model for the control.

        This is a wrapped instance for the model property.
        It add some additional properties and methods to the model.
        """
        if self._model_ex is None:
            # pylint: disable=import-outside-toplevel
            # pylint: disable=redefined-outer-name
            from ooodev.dialog.dl_control.model.model_checkbox import ModelCheckbox

            self._model_ex = ModelCheckbox(self.model)
        return self._model_ex

    @property
    def triple_state(self) -> bool:
        """
        Gets/Sets the triple state Same as ``tri_state`` property.

        Specifies if the checkbox control may appear dimmed (grayed) or not.
        """
        return self.tri_state

    @triple_state.setter
    def triple_state(self, value: bool) -> None:
        self.tri_state = value

    @property
    def view(self) -> UnoControlCheckBox:
        # pylint: disable=no-member
        return cast("UnoControlCheckBox", super().view)

    @property
    def view_ex(self) -> ViewCheckBox:
        """
        Gets the extended View for the control.

        This is a wrapped instance for the view property.
        It add some additional properties and methods to the view.
        """
        if self._view_ex is None:
            # pylint: disable=import-outside-toplevel
            # pylint: disable=redefined-outer-name
            from ooodev.dialog.dl_control.view.view_check_box import ViewCheckBox

            self._view_ex = ViewCheckBox(self.view)
        return self._view_ex

    # endregion Properties


if mock_g.FULL_IMPORT:
    from ooodev.dialog.dl_control.model.model_checkbox import ModelCheckbox
    from ooodev.dialog.dl_control.view.view_check_box import ViewCheckBox
