from __future__ import annotations
from typing import TYPE_CHECKING, cast
from pathlib import Path
import datetime
from ooodev.dialog import Dialogs, ImageScaleModeEnum, BorderKind
from ooodev.utils import lo as mLo
from ooodev.utils.gui import GUI
from ooodev.office.calc import Calc
from ooodev.utils.file_io import FileIO


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
        border_kind = BorderKind.BORDER_3D
        width = 800
        height = 700
        btn_width = 100
        btn_height = 30
        margin = 4
        box_height = 30
        if border_kind != BorderKind.BORDER_3D:
            padding = 10
        else:
            padding = 14

        ctl_lbl = Dialogs.insert_label(
            dialog_ctrl=dialog, label=msg, x=margin, y=margin, width=width - (margin * 2), height=20
        )
        sz = ctl_lbl.getPosSize()
        if is_password:
            txt_input = Dialogs.insert_password_field(
                dialog_ctrl=dialog,
                text=input_value,
                x=sz.X,
                y=sz.Height + sz.Y + 4,
                width=sz.Width,
                height=sz.Height,
                border=border_kind,
            )
        else:
            txt_input = Dialogs.insert_text_field(
                dialog_ctrl=dialog,
                text=input_value,
                x=sz.X,
                y=sz.Height + sz.Y + 4,
                width=sz.Width,
                height=sz.Height,
                border=border_kind,
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

        sz = txt_input.getPosSize()
        ctl_chk1 = Dialogs.insert_check_box(
            dialog_ctrl=dialog,
            label="Check Box 1",
            x=sz.X,
            y=sz.Height + sz.Y + padding,
            width=200,
            height=20,
            tri_state=False,
            state=Dialogs.StateEnum.CHECKED,
        )

        sz = ctl_chk1.getPosSize()
        ctl_chk2 = Dialogs.insert_check_box(
            dialog_ctrl=dialog,
            label="Check Box 2",
            x=sz.X,
            y=sz.Height + sz.Y,
            width=sz.Width,
            height=sz.Height,
            tri_state=False,
            state=Dialogs.StateEnum.NOT_CHECKED,
        )

        sz = ctl_chk2.getPosSize()
        ctl_chk3 = Dialogs.insert_check_box(
            dialog_ctrl=dialog,
            label="Check Box 3",
            x=sz.X,
            y=sz.Height + sz.Y,
            width=sz.Width,
            height=sz.Height,
            tri_state=True,
            state=Dialogs.StateEnum.DONT_KNOW,
        )

        sz = ctl_chk1.getPosSize()
        ctl_date = Dialogs.insert_date_field(
            dialog_ctrl=dialog,
            x=sz.Width + padding,
            y=sz.Y,
            width=190,
            height=box_height,
            date_value=datetime.datetime.now(),
            border=border_kind,
        )
        sz = ctl_date.getPosSize()
        ctl_currency = Dialogs.insert_currency_field(
            dialog_ctrl=dialog,
            x=sz.Width + sz.X + padding,
            y=sz.Y,
            width=sz.Width,
            height=sz.Height,
            value=123.45,
            spin_button=True,
            border=border_kind,
        )
        sz = ctl_currency.getPosSize()
        ctl_pattern = Dialogs.insert_pattern_field(
            dialog_ctrl=dialog,
            x=sz.X,
            y=sz.Y + sz.Height + padding,
            width=sz.Width,
            height=sz.Height,
            edit_mask="NNLNNLLLLL",
            literal_mask="__.__.2025",
            border=border_kind,
        )

        sz_date = ctl_date.getPosSize()
        ctl_num_field = Dialogs.insert_numeric_field(
            dialog_ctrl=dialog,
            x=sz_date.X,
            y=sz_date.Y + sz_date.Height + padding,
            width=sz.Width,
            height=box_height,
            value=123,
            spin_button=True,
            border=border_kind,
        )

        sz_numeric = ctl_num_field.getPosSize()
        # sz_fmt = ctl_formatted.getPosSize()
        ctl_combo1 = Dialogs.insert_combo_box(
            dialog_ctrl=dialog,
            x=margin,
            y=sz_numeric.Height + sz_numeric.Y + padding,
            width=200,
            height=box_height,
            entries=["Item 1", "Item 2", "Item 3"],
            border=border_kind,
        )

        sz = ctl_combo1.getPosSize()

        ctl_progress = Dialogs.insert_progress_bar(
            dialog_ctrl=dialog,
            x=sz_date.X,
            y=sz.Y,
            width=400,
            height=box_height,
            min=1,
            value=67,
            border=border_kind,
        )

        ctl_file = Dialogs.insert_file_control(
            dialog_ctrl=dialog,
            x=sz.X,
            y=sz.Height + sz.Y + padding,
            width=200,
            height=box_height,
            border=border_kind,
        )
        sz = ctl_file.getPosSize()
        ctl_ln = Dialogs.insert_fixed_line(
            dialog_ctrl=dialog,
            x=margin,
            y=sz.Height + sz.Y + padding,
            width=width - (margin * 2),
            height=1,
        )

        sz = ctl_ln.getPosSize()
        ctl_formatted = Dialogs.insert_formatted_field(
            dialog_ctrl=dialog,
            x=margin,
            y=sz.Height + sz.Y + padding,
            width=200,
            height=box_height,
            spin_button=True,
            value=3,
            border=border_kind,
        )

        sz = ctl_formatted.getPosSize()
        ctl_gb1 = Dialogs.insert_group_box(
            dialog_ctrl=dialog,
            x=margin,
            y=sz.Height + sz.Y + padding,
            width=round((width // 2) - ((padding * 2) * 0.75)),
            height=100,
            label="Group Box One",
        )

        # insert radio buttons into group box one
        sz = ctl_gb1.getPosSize()
        rb1 = Dialogs.insert_radio_button(
            dialog_ctrl=dialog,
            label="Radio Button 1",
            x=sz.X + padding,
            y=sz.Y + 10,
            width=sz.Width - (padding * 2),
            height=20,
        )
        rb_sz = rb1.getPosSize()
        for i in range(1, 4):
            _ = Dialogs.insert_radio_button(
                dialog_ctrl=dialog,
                label=f"Radio Button {i + 1}",
                x=rb_sz.X,
                y=rb_sz.Y + (rb_sz.Height * i),
                width=rb_sz.Width,
                height=rb_sz.Height,
            )

        sz = ctl_gb1.getPosSize()
        ctl_gb2 = Dialogs.insert_group_box(
            dialog_ctrl=dialog,
            x=sz.X + sz.Width + padding,
            y=sz.Y,
            width=sz.Width,
            height=sz.Height,
            label="Group Box Two",
        )

        # insert radio buttons into group box two
        sz = ctl_gb2.getPosSize()
        rb2 = Dialogs.insert_radio_button(
            dialog_ctrl=dialog,
            label="Radio Button 1",
            x=sz.X + padding,
            y=sz.Y + 10,
            width=sz.Width - (padding * 2),
            height=20,
        )
        rb_sz = rb2.getPosSize()
        for i in range(1, 4):
            _ = Dialogs.insert_radio_button(
                dialog_ctrl=dialog,
                label=f"Radio Button {i + 1}",
                x=rb_sz.X,
                y=rb_sz.Y + (rb_sz.Height * i),
                width=rb_sz.Width,
                height=rb_sz.Height,
            )

        sz = ctl_gb1.getPosSize()
        ctl_link = Dialogs.insert_hyperlink(
            dialog_ctrl=dialog,
            x=margin,
            y=sz.Height + sz.Y + padding,
            width=200,
            height=20,
            label="OOO Development Tools",
            url="https://python-ooo-dev-tools.readthedocs.io/en/latest/index.html",
            # Enabled=False,
        )
        sz = ctl_gb2.getPosSize()
        # file:///workspace/ooouno-dev-tools/tests/fixtures/image/img_brick.png
        pth = Path(__file__).parent.parent.parent / "fixtures" / "image" / "img_brick.png"
        ctl_img = Dialogs.insert_image_control(
            dialog_ctrl=dialog,
            x=sz.X,
            y=sz.Y + sz.Height + padding,
            width=120,
            height=120,
            image_url=FileIO.fnm_to_url(pth),
            scale=ImageScaleModeEnum.ANISOTROPIC,
            border=border_kind,
        )

        sz = ctl_img.getPosSize()
        ctl_list_box = Dialogs.insert_list_box(
            dialog_ctrl=dialog,
            x=sz.X + sz.Width + padding,
            y=sz.Y,
            width=sz.Width,
            entries=["Item 1", "Item 2", "Item 3", "Item 4", "Item 5"],
            drop_down=False,
            border=border_kind,
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


def main():
    with mLo.Lo.Loader(mLo.Lo.ConnectSocket(), opt=mLo.Lo.Options(verbose=True)):
        doc = Calc.create_doc()
        GUI.set_visible(visible=True, doc=doc)
        run()


def run() -> None:
    print(Input.get_input("title", "msg", "input_value"))


if __name__ == "__main__":
    main()
