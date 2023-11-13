from typing import TYPE_CHECKING, cast
from .dialogs import Dialogs
from ..utils import lo as mLo

from com.sun.star.awt import XControlModel
from com.sun.star.awt import XDialog

from ooo.dyn.awt.pos_size import PosSize
from ooo.dyn.awt.push_button_type import PushButtonType

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlDialog


class Input:
    @staticmethod
    def get_input(
        title: str,
        msg: str,
        input_value: str = "",
        ok_lbl: str = "OK",
        cancel_lbl: str = "Cancel",
        is_password: bool = False,
    ) -> str:
        """
        Displays an input box and returns the results.

        Args:
            title (str): Title for the dialog
            msg (str): Message to display such as "Input your Name"
            input_value (str, optional): Value of input box when first displayed.
            ok_lbl (str, optional): OK button Label. Defaults to "OK".
            cancel_lbl (str, optional): Cancel Button Label. Defaults to "Cancel".
            is_password (bool, optional): Determines if the input box is masked for password input. Defaults to False.

        Returns:
            str: The value of input or empty string.
        """
        dialog = cast(
            "UnoControlDialog",
            mLo.Lo.create_instance_mcf(XDialog, "com.sun.star.awt.UnoControlDialog", raise_err=True),
        )
        dialog_model = mLo.Lo.create_instance_mcf(
            XControlModel, "com.sun.star.awt.UnoControlDialogModel", raise_err=True
        )

        dialog.setModel(dialog_model)
        width = 500
        height = 80
        btn_width = 80
        btn_height = 20
        margin = 4

        ctl_lbl = Dialogs.insert_label(
            dialog_ctrl=dialog, label=msg, x=margin, y=margin, width=width - (margin * 2), height=20
        )
        sz = ctl_lbl.getPosSize()
        if is_password:
            txt_input = Dialogs.insert_password_field(
                dialog_ctrl=dialog, text=input_value, x=sz.X, y=sz.Height + sz.Y + 4, width=sz.Width, height=sz.Height
            )
        else:
            txt_input = Dialogs.insert_text_field(
                dialog_ctrl=dialog, text=input_value, x=sz.X, y=sz.Height + sz.Y + 4, width=sz.Width, height=sz.Height
            )
        ctl_btn_cancel = Dialogs.insert_button(
            dialog_ctrl=dialog,
            label=cancel_lbl,
            x=width - btn_width - margin,
            y=height - btn_height - margin,
            width=btn_width,
            height=btn_height,
            btn_type=PushButtonType.CANCEL,
        )
        sz = ctl_btn_cancel.getPosSize()
        _ = Dialogs.insert_button(
            dialog_ctrl=dialog,
            label=ok_lbl,
            x=sz.X - sz.Width - margin,
            y=sz.Y,
            width=btn_width,
            height=btn_height,
            btn_type=PushButtonType.OK,
            DefaultButton=True,
        )
        window = mLo.Lo.get_frame().getContainerWindow()
        ps = window.getPosSize()
        x = round(ps.Width / 2 - width / 2)
        y = round(ps.Height / 2 - height / 2)
        dialog.setTitle(title)
        dialog.setPosSize(x, y, width, height, PosSize.POSSIZE)
        dialog.setVisible(True)
        ret = txt_input.getModel().Text if dialog.execute() else ""  # type: ignore
        dialog.dispose()
        return ret
