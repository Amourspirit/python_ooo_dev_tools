from typing import TYPE_CHECKING, cast
from ..utils.dialogs import Dialogs
from ..utils import lo as mLo
from ..utils import sys_info as mSysInfo


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
        dialog = cast("UnoControlDialog", mLo.Lo.create_instance_mcf(XDialog, "com.sun.star.awt.UnoControlDialog"))
        dialog_model = mLo.Lo.create_instance_mcf(XControlModel, "com.sun.star.awt.UnoControlDialogModel")

        dialog.setModel(dialog_model)
        platform = mSysInfo.SysInfo.get_platform()
        if platform == mSysInfo.SysInfo.PlatformEnum.WINDOWS:
            width = 420
            height = 130
        else:
            width = 500
            height = 140
        btn_width = 50
        btn_height = 18

        _ = Dialogs.insert_label(dialog_ctrl=dialog, label=msg, x=4, y=4, width=200, height=8)
        if is_password:
            txt_input = Dialogs.insert_password_field(
                dialog_ctrl=dialog, text=input_value, x=4, y=18, width=200, height=12
            )
        else:
            txt_input = Dialogs.insert_text_field(
                dialog_ctrl=dialog, text=input_value, x=4, y=18, width=200, height=12
            )
        _ = Dialogs.insert_button(
            dialog_ctrl=dialog,
            label=ok_lbl,
            x=98,
            y=40,
            width=btn_width,
            height=btn_height,
            btn_type=PushButtonType.OK,
            DefaultButton=True,
        )
        _ = Dialogs.insert_button(
            dialog_ctrl=dialog,
            label=cancel_lbl,
            x=152,
            y=40,
            width=btn_width,
            height=btn_height,
            btn_type=PushButtonType.CANCEL,
        )
        window = mLo.Lo.get_frame().getContainerWindow()
        ps = window.getPosSize()
        x = ps.Width / 2 - width / 2
        y = ps.Height / 2 - height / 2
        dialog.setTitle(title)
        dialog.setPosSize(x, y, width, height, PosSize.POSSIZE)
        dialog.setVisible(True)
        ret = txt_input.getModel().Text if dialog.execute() else ""
        dialog.dispose()
        return ret
