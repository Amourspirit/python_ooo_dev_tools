from __future__ import annotations
from typing import TYPE_CHECKING, cast
from ooodev.dialog import Dialogs, BorderKind
from ooodev.utils import lo as mLo
from ooodev.utils.gui import GUI
from ooodev.office.calc import Calc


from com.sun.star.awt import XControlModel
from com.sun.star.awt import XDialog

from ooo.dyn.awt.pos_size import PosSize
from ooo.dyn.awt.push_button_type import PushButtonType

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlDialog
    from com.sun.star.awt import UnoControlDialogModel


class Tabs:
    def __init__(self) -> None:
        self._border_kind = BorderKind.BORDER_3D
        self._width = 600
        self._height = 500
        self._btn_width = 100
        self._btn_height = 30
        self._margin = 4
        self._box_height = 30
        self._title = "Tabs"
        self._msg = "Hello World!"
        if self._border_kind != BorderKind.BORDER_3D:
            self._padding = 10
        else:
            self._padding = 14
        self._init_dialog()

    def _init_dialog(self) -> None:
        self._dialog = cast(
            "UnoControlDialog",
            mLo.Lo.create_instance_mcf(XDialog, "com.sun.star.awt.UnoControlDialog", raise_err=True),
        )
        self._dialog_model = cast(
            "UnoControlDialogModel",
            mLo.Lo.create_instance_mcf(XControlModel, "com.sun.star.awt.UnoControlDialogModel", raise_err=True),
        )
        self._dialog.setModel(self._dialog_model)
        self._window = mLo.Lo.get_frame().getContainerWindow()
        ps = self._window.getPosSize()
        x = round(ps.Width / 2 - self._width / 2)
        y = round(ps.Height / 2 - self._height / 2)
        self._dialog.setPosSize(x, y, self._width, self._height, PosSize.POSSIZE)
        self._dialog.setTitle(self._title)
        self._dialog.setVisible(False)
        # createPeer() must be call before inserting tabs
        self._dialog.createPeer(self._dialog_model.createInstance("com.sun.star.awt.Toolkit"), self._window)  # type: ignore
        self._init_tab_control()

    def _init_tab_control(self) -> None:
        self._ctl_tab = Dialogs.insert_tab_control(
            dialog_ctrl=self._dialog,
            x=self._margin,
            y=self._margin,
            width=self._width - (self._margin * 2),
            height=self._height - (self._margin * 2),
        )
        self._init_tab_main()
        self._init_tab_oth()

    def _init_tab_main(self) -> None:
        self._tab_main = Dialogs.insert_tab_page(
            dialog_ctrl=self._dialog,
            tab_ctrl=self._ctl_tab,
            title="Main",
            tab_position=1,
        )
        tab_sz = self._tab_main.getPosSize()
        ctl_main_lbl = Dialogs.insert_label(
            dialog_ctrl=self._tab_main,
            label=self._msg,
            x=tab_sz.X + self._padding,
            y=tab_sz.Y + self._padding,
            width=100,
            height=20,
        )

    def _init_tab_oth(self) -> None:
        self._tab_oth = Dialogs.insert_tab_page(
            dialog_ctrl=self._dialog,
            tab_ctrl=self._ctl_tab,
            title="Other",
            tab_position=2,
        )
        tab_sz = self._tab_oth.getPosSize()
        ctl_oth_lbl = Dialogs.insert_label(
            dialog_ctrl=self._tab_oth,
            label="Nice Day!",
            x=tab_sz.X + self._padding,
            y=tab_sz.Y + self._padding,
            width=100,
            height=20,
        )

    def show(self) -> int:
        self._ctl_tab.ActiveTabPageID = 1
        self._dialog.setVisible(True)
        result = self._dialog.execute()
        self._dialog.dispose()
        return result


def main():
    with mLo.Lo.Loader(mLo.Lo.ConnectSocket(), opt=mLo.Lo.Options(verbose=True)):
        doc = Calc.create_doc()
        GUI.set_visible(visible=True, doc=doc)
        run()


def run() -> None:
    Tabs().show()


if __name__ == "__main__":
    main()
