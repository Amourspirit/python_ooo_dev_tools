from __future__ import annotations
from typing import Any, TYPE_CHECKING, cast
import uno  # pylint: disable=unused-import
from com.sun.star.awt import XControlModel
from com.sun.star.awt import XDialog

from ooo.dyn.awt.pos_size import PosSize

from ooodev.dialog import Dialogs, BorderKind, OrientationKind
from ooodev.events.args.event_args import EventArgs
from ooodev.calc import CalcDoc
from ooodev.loader import lo as mLo
from ooodev.utils.gui import GUI
from ooodev.utils.table_helper import TableHelper


if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlDialog
    from com.sun.star.awt import UnoControlDialogModel
    from com.sun.star.awt.tab import TabPageActivatedEvent
    from com.sun.star.awt.grid import GridSelectionEvent


class Tabs:
    # pylint: disable=unused-argument
    def __init__(self, doc: CalcDoc) -> None:
        self._doc = doc
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
        self._window = mLo.Lo.get_frame().getContainerWindow()
        ps = self._window.getPosSize()
        x = round(ps.Width / 2 - self._width / 2)
        y = round(ps.Height / 2 - self._height / 2)

        # self._dialog = cast(
        #     "UnoControlDialog",
        #     mLo.Lo.create_instance_mcf(XDialog, "com.sun.star.awt.UnoControlDialog", raise_err=True),
        # )
        self._dialog = self._doc.create_dialog(x=x, y=y, width=self._width, height=self._height, title=self._title)
        self._dialog.create_peer()

        # self._dialog_model = cast(
        #     "UnoControlDialogModel",
        #     mLo.Lo.create_instance_mcf(XControlModel, "com.sun.star.awt.UnoControlDialogModel", raise_err=True),
        # )
        # self._dialog.setModel(self._dialog_model)
        mLo.Lo.delay(300)  # wait for window to be created
        # self._doc.msgbox("Dialog created")
        self._dialog.set_pos_size(x, y, self._width, self._height, PosSize.POSSIZE)
        self._dialog.set_title(self._title)
        self._dialog.set_visible(False)

        # createPeer() must be call before inserting tabs
        # tk = self._dialog_model.createInstance("com.sun.star.awt.Toolkit")
        # self._dialog.createPeer(tk, self._window)  # type: ignore

        # tab offset will vary depending on border kind and Operating System
        self._tab_offset_vert = (self._margin * 3) + 30
        self._init_tab_control()
        self._active_page_page_id = 1

    def _init_tab_control(self) -> None:
        self._ctl_tab = self._dialog.insert_tab_control(
            x=self._margin,
            y=self._margin,
            width=self._width - (self._margin * 2),
            height=self._height - (self._margin * 2),
        )
        self._ctl_tab.add_event_tab_page_activated(self._fn_tab_activated)
        self._init_tab_tree()
        self._init_tab_table()
        self._init_tab_scroll_bar()
        self._init_tab_main()
        self._init_tab_oth()

    def _init_tab_main(self) -> None:
        self._tab_count += 1
        self._tab_main = self._dialog.insert_tab_page(
            tab_ctrl=self._ctl_tab,
            title="Main",
            tab_position=self._tab_count,
        )
        tab_sz = self._ctl_tab.view.getPosSize()
        ctl_main_lbl = self._dialog.insert_label(
            label=self._msg,
            x=tab_sz.X + self._padding,
            y=tab_sz.Y + self._padding,
            width=100,
            height=20,
        )

    def _init_tab_tree(self) -> None:
        self._tab_count += 1
        self._tab_tree = self._dialog.insert_tab_page(
            tab_ctrl=self._ctl_tab,
            title="Tree",
            tab_position=self._tab_count,
        )
        tab_sz = self._ctl_tab.view.getPosSize()
        self._tree1 = self._dialog.insert_tree_control(
            x=tab_sz.X + self._padding,
            y=tab_sz.Y + self._padding,
            width=tab_sz.Width - (self._padding * 2),
            height=300,
            border=self._border_kind,
            dialog_ctrl=self._tab_tree.view,
        )
        root1 = self._tree1.create_root(display_value="Root 1")
        if root1:
            _ = self._tree1.add_sub_node(parent_node=root1, display_value="Node 1")
            _ = self._tree1.add_sub_node(parent_node=root1, display_value="Node 3")
            _ = self._tree1.add_sub_node(parent_node=root1, display_value="Node 3")
        root2 = self._tree1.create_root(display_value="Root 2")
        if root2:
            _ = self._tree1.add_sub_node(parent_node=root2, display_value="Node 1")
            _ = self._tree1.add_sub_node(parent_node=root2, display_value="Node 3")
            _ = self._tree1.add_sub_node(parent_node=root2, display_value="Node 3")

    def _init_tab_table(self) -> None:
        self._tab_count += 1
        self._tab_table = self._dialog.insert_tab_page(
            tab_ctrl=self._ctl_tab,
            title="Table",
            tab_position=self._tab_count,
        )
        tab_sz = self._ctl_tab.view.getPosSize()
        self._ctl_table1 = self._dialog.insert_table_control(
            x=tab_sz.X + self._padding,
            y=tab_sz.Y + self._padding,
            width=tab_sz.Width - (self._padding * 2),
            height=300,
            grid_lines=True,
            col_header=True,
            row_header=True,
            dialog_ctrl=self._tab_table.view,
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
        self._ctl_table1.set_table_data(
            data=tbl,
            align="RLC",
            widths=widths,
            has_row_headers=has_row_headers,
            has_colum_headers=True,
        )
        self._ctl_table1.add_event_selection_changed(self._fn_grid_selection_changed)
        # test that grid clears and adds new data
        # widths[0] = 100
        # Dialogs.set_table_data(table=ctl_table1, data=tbl, align="RLC", widths=widths)

    def _init_tab_scroll_bar(self) -> None:
        self._tab_count += 1
        self._tab_scroll_bar = self._dialog.insert_tab_page(
            tab_ctrl=self._ctl_tab,
            title="Scroll",
            tab_position=self._tab_count,
        )
        tab_sz = self._ctl_tab.view.getPosSize()
        ctl_scroll_bar1 = self._dialog.insert_scroll_bar(
            # x=tab_sz.X + self._padding,
            x=0,
            y=0,
            width=tab_sz.Width - 22,
            height=20,
            border=self._border_kind,
            dialog_ctrl=self._tab_scroll_bar.view,
        )

        ctl_scroll_bar2 = self._dialog.insert_scroll_bar(
            x=tab_sz.Width - 22,
            # y=tab_sz.Y,
            y=0,
            width=20,
            height=tab_sz.Height - self._tab_offset_vert,
            border=self._border_kind,
            orientation=OrientationKind.VERTICAL,
            dialog_ctrl=self._tab_scroll_bar.view,
        )

    def _init_tab_oth(self) -> None:
        self._tab_count += 1
        self._tab_oth = self._dialog.insert_tab_page(
            tab_ctrl=self._ctl_tab,
            title="Other",
            tab_position=self._tab_count,
        )
        tab_sz = self._tab_oth.view.getPosSize()
        ctl_oth_lbl = self._dialog.insert_label(
            label="Nice Day!",
            x=tab_sz.X + self._padding,
            y=tab_sz.Y + self._padding,
            width=100,
            height=20,
            dialog_ctrl=self._tab_oth.view,
        )
        self._ctl_combo1 = self._dialog.insert_combo_box(
            x=tab_sz.X + self._padding,
            y=100,
            width=200,
            height=self._box_height,
            entries=["Item 1", "Item 2", "Item 3"],
            dialog_ctrl=self._tab_oth.view,
        )

    def show(self) -> int:
        self._ctl_tab.active_tab_page_id = self._active_page_page_id
        self._dialog.set_visible(True)
        result = self._dialog.execute()
        self._dialog.dispose()
        return result

    # region Event Handlers
    def _init_handlers(self) -> None:
        def on_tab_activated(src: Any, event: EventArgs, control_src: Any, *args, **kwargs):
            self.on_tab_activated(src, event, control_src, *args, **kwargs)

        def on_grid_selection_changed(src: Any, event: EventArgs, control_src: Any, *args, **kwargs) -> None:
            self.on_grid_selection_changed(src, event, control_src, *args, **kwargs)

        self._fn_tab_activated = on_tab_activated
        self._fn_grid_selection_changed = on_grid_selection_changed

    def on_tab_activated(self, src: Any, event: EventArgs, control_src: Any, *args, **kwargs) -> None:
        print("Tab Changed:", control_src.name)
        itm_event = cast("TabPageActivatedEvent", event.event_data)
        self._active_page_page_id = itm_event.TabPageID
        print("Active ID:", self._active_page_page_id)

    def on_grid_selection_changed(self, src: Any, event: EventArgs, control_src: Any, *args, **kwargs) -> None:
        print("Grid Selection Changed:", control_src.name)
        itm_event = cast("GridSelectionEvent", event.event_data)
        print("Selected row indexes:", itm_event.SelectedRowIndexes)
        print("Selected column indexes:", itm_event.SelectedColumnIndexes)

    # endregion Event Handlers


def main():
    with mLo.Lo.Loader(mLo.Lo.ConnectSocket(), opt=mLo.Lo.Options(verbose=True)):
        doc = CalcDoc.create_doc(visible=True)
        # doc.activate()
        run(doc)


def run(doc: CalcDoc) -> None:
    tabs = Tabs(doc)
    tabs.show()


if __name__ == "__main__":
    main()
