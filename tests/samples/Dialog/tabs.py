from __future__ import annotations
from typing import Any, TYPE_CHECKING, cast
from ooodev.dialog import Dialogs, BorderKind, OrientationKind, HorzVertKind
from ooodev.events.args.event_args import EventArgs
from ooodev.office.calc import Calc
from ooodev.utils import lo as mLo
from ooodev.utils.gui import GUI
from ooodev.utils.table_helper import TableHelper


from com.sun.star.awt import XControlModel
from com.sun.star.awt import XDialog

from ooo.dyn.awt.pos_size import PosSize
from ooo.dyn.awt.push_button_type import PushButtonType

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlDialog
    from com.sun.star.awt import UnoControlDialogModel
    from com.sun.star.awt.tab import TabPageActivatedEvent


class Tabs:
    def __init__(self) -> None:
        self._border_kind = BorderKind.BORDER_SIMPLE
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
        self._tab_count = 0
        self._init_dialog()

    def _init_dialog(self) -> None:
        self._init_handlers()

        self._dialog = cast(
            "UnoControlDialog",
            mLo.Lo.create_instance_mcf(XDialog, "com.sun.star.awt.UnoControlDialog", raise_err=True),
        )
        self._dialog_model = cast(
            "UnoControlDialogModel",
            mLo.Lo.create_instance_mcf(XControlModel, "com.sun.star.awt.UnoControlDialogModel", raise_err=True),
        )
        self._dialog.setModel(self._dialog_model)
        mLo.Lo.delay(300)  # wait for window to be created
        self._window = mLo.Lo.get_frame().getContainerWindow()
        ps = self._window.getPosSize()
        x = round(ps.Width / 2 - self._width / 2)
        y = round(ps.Height / 2 - self._height / 2)
        self._dialog.setPosSize(x, y, self._width, self._height, PosSize.POSSIZE)
        self._dialog.setTitle(self._title)
        self._dialog.setVisible(False)

        # createPeer() must be call before inserting tabs
        self._dialog.createPeer(self._dialog_model.createInstance("com.sun.star.awt.Toolkit"), self._window)  # type: ignore

        # tab offset will vary depending on border kind and Operating System
        self._tab_offset_vert = (self._margin * 3) + 30
        self._init_tab_control()
        self._active_page_page_id = 1

    def _init_tab_control(self) -> None:
        self._ctl_tab = Dialogs.insert_tab_control(
            dialog_ctrl=self._dialog,
            x=self._margin,
            y=self._margin,
            width=self._width - (self._margin * 2),
            height=self._height - (self._margin * 2),
        )
        self._ctl_tab.add_event_tab_page_activated(self._fn_tab_activated)
        self._init_tab_table()
        self._init_tab_scroll_bar()
        self._init_tab_main()
        self._init_tab_oth()

    def _init_tab_main(self) -> None:
        self._tab_count += 1
        self._tab_main = Dialogs.insert_tab_page(
            dialog_ctrl=self._dialog,
            tab_ctrl=self._ctl_tab,
            title="Main",
            tab_position=self._tab_count,
        )
        tab_sz = self._ctl_tab.view.getPosSize()
        ctl_main_lbl = Dialogs.insert_label(
            dialog_ctrl=self._tab_main.view,
            label=self._msg,
            x=tab_sz.X + self._padding,
            y=tab_sz.Y + self._padding,
            width=100,
            height=20,
        )

    def _init_tab_table(self) -> None:
        self._tab_count += 1
        self._tab_table = Dialogs.insert_tab_page(
            dialog_ctrl=self._dialog,
            tab_ctrl=self._ctl_tab,
            title="Table",
            tab_position=self._tab_count,
        )
        tab_sz = self._ctl_tab.view.getPosSize()
        ctl_table1 = Dialogs.insert_table_control(
            dialog_ctrl=self._tab_table.view,
            x=tab_sz.X + self._padding,
            y=tab_sz.Y + self._padding,
            width=tab_sz.Width - (self._padding * 2),
            height=300,
            grid_lines=True,
            col_header=True,
            row_header=True,
        )
        num_cols = 5
        num_rows = 10
        tbl = TableHelper.make_2d_array(num_rows=num_rows, num_cols=num_cols)
        has_row_headers = False
        headers = [f"Column {i if has_row_headers else i + 1}" for i in range(num_cols)]
        widths = [220 for _ in range(num_cols)]
        widths[0] = 200
        # set_table_data() will handel to many or to few widths
        # widths.pop()
        # widths.append(100)
        tbl.insert(0, headers)
        Dialogs.set_table_data(
            table=ctl_table1,
            data=tbl,
            align="RLC",
            widths=widths,
            has_row_headers=has_row_headers,
            has_colum_headers=True,
        )
        # test that grid clears and adds new data
        # widths[0] = 100
        # Dialogs.set_table_data(table=ctl_table1, data=tbl, align="RLC", widths=widths)

    def _init_tab_scroll_bar(self) -> None:
        self._tab_count += 1
        self._tab_scroll_bar = Dialogs.insert_tab_page(
            dialog_ctrl=self._dialog,
            tab_ctrl=self._ctl_tab,
            title="Scroll",
            tab_position=self._tab_count,
        )
        tab_sz = self._ctl_tab.view.getPosSize()
        ctl_scroll_bar1 = Dialogs.insert_scroll_bar(
            dialog_ctrl=self._tab_scroll_bar.view,
            # x=tab_sz.X + self._padding,
            x=0,
            y=0,
            width=tab_sz.Width - 22,
            height=20,
            border=self._border_kind,
        )

        ctl_scroll_bar2 = Dialogs.insert_scroll_bar(
            dialog_ctrl=self._tab_scroll_bar.view,
            x=tab_sz.Width - 22,
            # y=tab_sz.Y,
            y=0,
            width=20,
            height=tab_sz.Height - self._tab_offset_vert,
            border=self._border_kind,
            orientation=OrientationKind.VERTICAL,
        )

    def _init_tab_oth(self) -> None:
        self._tab_count += 1
        self._tab_oth = Dialogs.insert_tab_page(
            dialog_ctrl=self._dialog,
            tab_ctrl=self._ctl_tab,
            title="Other",
            tab_position=self._tab_count,
        )
        tab_sz = self._tab_oth.view.getPosSize()
        ctl_oth_lbl = Dialogs.insert_label(
            dialog_ctrl=self._tab_oth.view,
            label="Nice Day!",
            x=tab_sz.X + self._padding,
            y=tab_sz.Y + self._padding,
            width=100,
            height=20,
        )

    def show(self) -> int:
        self._ctl_tab.active_tab_page_id = self._active_page_page_id
        self._dialog.setVisible(True)
        result = self._dialog.execute()
        self._dialog.dispose()
        return result

    # region Event Handlers
    def _init_handlers(self) -> None:
        def on_tab_activated(src: Any, event: EventArgs, control_src: Any, *args, **kwargs):
            self.on_tab_activated(src, event, control_src, *args, **kwargs)

        self._fn_tab_activated = on_tab_activated

    def on_tab_activated(self, src: Any, event: EventArgs, control_src: Any, *args, **kwargs) -> None:
        print("Tab Changed:", control_src.name)
        itm_event = cast("TabPageActivatedEvent", event.event_data)
        self._active_page_page_id = itm_event.TabPageID
        print("Active ID:", self._active_page_page_id)

    # endregion Event Handlers


def main():
    with mLo.Lo.Loader(mLo.Lo.ConnectSocket(), opt=mLo.Lo.Options(verbose=True)):
        doc = Calc.create_doc()
        GUI.set_visible(visible=True, doc=doc)
        run()


def run() -> None:
    Tabs().show()


if __name__ == "__main__":
    main()
