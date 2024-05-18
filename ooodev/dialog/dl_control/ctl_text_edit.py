# region imports
from __future__ import annotations
import contextlib
from typing import Any, cast, TYPE_CHECKING
import os
import uno  # pylint: disable=unused-import

# com.sun.star.awt.Selection
from ooo.dyn.awt.selection import Selection

# pylint: disable=useless-import-alias
from ooodev.mock import mock_g
from ooodev.adapter.awt.text_events import TextEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind
from ooodev.adapter.awt.uno_control_edit_model_partial import UnoControlEditModelPartial
from ooodev.dialog.dl_control.ctl_base import DialogControlBase, _create_control

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlEdit  # service
    from com.sun.star.awt import UnoControlEditModel  # service
    from com.sun.star.awt import XWindowPeer
    from ooodev.dialog.dl_control.model.model_text_edit import ModelTextEdit
    from ooodev.dialog.dl_control.view.view_text_edit import ViewTextEdit
# endregion imports


class CtlTextEdit(DialogControlBase, UnoControlEditModelPartial, TextEvents):
    """Class for Text Edit Control"""

    # pylint: disable=unused-argument

    # region init
    def __init__(self, ctl: UnoControlEdit) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlEdit): Button Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        DialogControlBase.__init__(self, ctl)
        UnoControlEditModelPartial.__init__(self, component=self.get_model())
        generic_args = self._get_generic_args()
        # EventArgs.event_data will contain the ActionEvent
        TextEvents.__init__(self, trigger_args=generic_args, cb=self._on_text_events_listener_add_remove)
        self._model_ex = None
        self._view_ex = None

    # endregion init

    def __repr__(self) -> str:
        if hasattr(self, "name"):
            return f"CtlTextEdit({self.name})"
        return "CtlTextEdit"

    # region Lazy Listeners
    def _on_text_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addTextListener(self.events_listener_text)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides

    def get_view_ctl(self) -> UnoControlEdit:
        return cast("UnoControlEdit", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlEdit``"""
        return "com.sun.star.awt.UnoControlEdit"

    def get_model(self) -> UnoControlEditModel:
        """Gets the Model for the control"""
        return cast("UnoControlEditModel", self.get_view_ctl().getModel())

    def get_control_kind(self) -> DialogControlKind:
        """Gets the control kind. Returns ``DialogControlKind.EDIT``"""
        return DialogControlKind.EDIT

    def get_control_named_kind(self) -> DialogControlNamedKind:
        """Gets the control named kind. Returns ``DialogControlNamedKind.EDIT``"""
        return DialogControlNamedKind.EDIT

    # endregion Overrides

    # region Text Methods
    def write_line(self, line: str = "") -> bool:
        """
        Add a new line to a multi-line text control

        Args:
            line (str, optional): Specifies a line to insert at the end of the text box
                a newline character will be inserted before the line, if relevant.

        Returns:
            bool: True if successful, False otherwise
        """
        # if not self.model.MultiLine:
        #     return False
        with contextlib.suppress(Exception):
            # will raise an exception if not multi-line
            self.model.HardLineBreaks = True
            sel = Selection()
            text_len = len(self.text)
            if text_len == 0:
                sel.Min = 0
                sel.Max = 0
                self.text = line
            else:
                # Put cursor at the end of the actual text
                sel.Min = text_len
                sel.Max = text_len
                self.view.insertText(sel, f"{os.linesep}{line}")
            # Put the cursor at the end of the inserted text
            sel.Max += len(os.linesep) + len(line)
            sel.Min = sel.Max
            self.view.setSelection(sel)
            return True
        return False

    # endregion Text Methods

    # region Static Methods
    @staticmethod
    def create(win: XWindowPeer, **kwargs: Any) -> "CtlTextEdit":
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
            CtlTextEdit: New instance of the control.

        Note:
            The `UnoControlDialogElement <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1awt_1_1UnoControlDialogElement.html>`__
            interface is not included when creating the control with a window peer.
        """
        ctrl = _create_control("com.sun.star.awt.UnoControlEditModel", win, **kwargs)
        return CtlTextEdit(ctl=ctrl)

    # endregion Static Methods

    # region Properties

    @property
    def model(self) -> UnoControlEditModel:
        # pylint: disable=no-member
        return cast("UnoControlEditModel", super().model)

    @property
    def model_ex(self) -> ModelTextEdit:
        """
        Gets the extended Model for the control.

        This is a wrapped instance for the model property.
        It add some additional properties and methods to the model.
        """
        # pylint: disable=no-member
        if self._model_ex is None:
            # pylint: disable=import-outside-toplevel
            # pylint: disable=redefined-outer-name
            from ooodev.dialog.dl_control.model.model_text_edit import ModelTextEdit

            self._model_ex = ModelTextEdit(self.model)
        return self._model_ex

    @property
    def view(self) -> UnoControlEdit:
        # pylint: disable=no-member
        return cast("UnoControlEdit", super().view)

    @property
    def view_ex(self) -> ViewTextEdit:
        """
        Gets the extended View for the control.

        This is a wrapped instance for the view property.
        It add some additional properties and methods to the view.
        """
        # pylint: disable=no-member
        if self._view_ex is None:
            # pylint: disable=import-outside-toplevel
            # pylint: disable=redefined-outer-name
            from ooodev.dialog.dl_control.view.view_text_edit import ViewTextEdit

            self._view_ex = ViewTextEdit(self.view)
        return self._view_ex

    # endregion Properties


if mock_g.FULL_IMPORT:
    from ooodev.dialog.dl_control.model.model_text_edit import ModelTextEdit
    from ooodev.dialog.dl_control.view.view_text_edit import ViewTextEdit
