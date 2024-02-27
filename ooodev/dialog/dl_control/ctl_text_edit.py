# region imports
from __future__ import annotations
import contextlib
from typing import Any, cast, TYPE_CHECKING
import os
import uno  # pylint: disable=unused-import

# com.sun.star.awt.Selection
from ooo.dyn.awt.selection import Selection

# pylint: disable=useless-import-alias
from ooodev.adapter.awt.text_events import TextEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind
from ooodev.adapter.awt.uno_control_edit_model_partial import UnoControlEditModelPartial

from ooodev.dialog.dl_control.ctl_base import DialogControlBase

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlEdit  # service
    from com.sun.star.awt import UnoControlEditModel  # service
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
        UnoControlEditModelPartial.__init__(self)
        generic_args = self._get_generic_args()
        # EventArgs.event_data will contain the ActionEvent
        TextEvents.__init__(self, trigger_args=generic_args, cb=self._on_text_events_listener_add_remove)

    # endregion init

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

    # region Properties

    @property
    def model(self) -> UnoControlEditModel:
        # pylint: disable=no-member
        return cast("UnoControlEditModel", super().model)

    @property
    def view(self) -> UnoControlEdit:
        # pylint: disable=no-member
        return cast("UnoControlEdit", super().view)

    # endregion Properties


# ctl = CtlButton(None)
