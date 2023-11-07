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
    from com.sun.star.awt import UnoControlDialogModel


class Tabs:
    @staticmethod
    def show() -> None:
        border_kind = BorderKind.BORDER_3D
        width = 600
        height = 500
        btn_width = 100
        btn_height = 30
        margin = 4
        box_height = 30
        title = "Tabs"
        msg = "Hello World!"
        if border_kind != BorderKind.BORDER_3D:
            padding = 10
        else:
            padding = 14
        dialog = cast(
            "UnoControlDialog",
            mLo.Lo.create_instance_mcf(XDialog, "com.sun.star.awt.UnoControlDialog", raise_err=True),
        )
        dialog_model = cast(
            "UnoControlDialogModel",
            mLo.Lo.create_instance_mcf(XControlModel, "com.sun.star.awt.UnoControlDialogModel", raise_err=True),
        )

        dialog.setModel(dialog_model)
        window = mLo.Lo.get_frame().getContainerWindow()
        ps = window.getPosSize()
        x = round(ps.Width / 2 - width / 2)
        y = round(ps.Height / 2 - height / 2)
        # dialog.setPosSize(-10, -10, 1, 1, PosSize.POSSIZE)
        dialog.setPosSize(x, y, width, height, PosSize.POSSIZE)
        dialog.setTitle(title)
        dialog.setVisible(False)
        # createPeer() must be call before inserting tabs
        dialog.createPeer(dialog_model.createInstance("com.sun.star.awt.Toolkit"), window)  # type: ignore

        ctl_tab = Dialogs.insert_tab_control(
            dialog_ctrl=dialog,
            x=margin,
            y=margin,
            width=width - (margin * 2),
            height=height - (margin * 2),
        )

        tab_main = Dialogs.insert_tab_page(
            dialog_ctrl=dialog,
            tab_ctrl=ctl_tab,
            title="Main",
            tab_position=1,
        )
        tab_oth = Dialogs.insert_tab_page(
            dialog_ctrl=dialog,
            tab_ctrl=ctl_tab,
            title="Other",
            tab_position=2,
        )
        tab_sz = tab_main.getPosSize()
        ctl_main_lbl = Dialogs.insert_label(
            dialog_ctrl=tab_main,
            label=msg,
            x=tab_sz.X + padding,
            y=tab_sz.Y + padding,
            width=100,
            height=20,
        )
        ctl_oth_lbl = Dialogs.insert_label(
            dialog_ctrl=tab_oth,
            label="Nice Day!",
            x=tab_sz.X + padding,
            y=tab_sz.Y + padding,
            width=100,
            height=20,
        )

        ctl_tab.ActiveTabPageID = 1

        # dialog.setPosSize(x, y, width, height, PosSize.POSSIZE)
        dialog.setVisible(True)
        dialog.execute()
        dialog.dispose()


def main():
    with mLo.Lo.Loader(mLo.Lo.ConnectSocket(), opt=mLo.Lo.Options(verbose=True)):
        doc = Calc.create_doc()
        GUI.set_visible(visible=True, doc=doc)
        run()


def run() -> None:
    Tabs.show()


if __name__ == "__main__":
    main()
