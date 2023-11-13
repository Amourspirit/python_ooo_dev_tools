from __future__ import annotations
import datetime
from typing import Any, TYPE_CHECKING, cast
from pathlib import Path
import datetime
from ooodev.dialog import Dialogs, ImageScaleModeEnum, BorderKind, DateFormatKind
from ooodev.utils import lo as mLo
from ooodev.utils.gui import GUI
from ooodev.office.calc import Calc
from ooodev.events.args.event_args import EventArgs

from ooo.dyn.awt.pos_size import PosSize


from ooodev.dialog.search.tree_search.search_tree import SearchTree
from ooodev.dialog.search.tree_search.rule_data_insensitive import RuleDataInsensitive
from ooodev.dialog.search.tree_search.rule_data_sensitive import RuleDataSensitive
from ooodev.dialog.search.tree_search.rule_text_insensitive import RuleTextInsensitive


if TYPE_CHECKING:
    from ooodev.dialog.dl_control.ctl_tree import CtlTree


class Tree:
    def __init__(self) -> None:
        self._border_kind = BorderKind.NONE
        self._title = "Tree Example"
        self._width = 800
        self._height = 700
        self._btn_width = 100
        self._btn_height = 30
        self._margin = 4
        self._box_height = 30
        if self._border_kind != BorderKind.BORDER_3D:
            self._padding = 10
        else:
            self._padding = 14

        self._init_handlers()

        self._dialog = Dialogs.create_dialog(
            x=-1,
            y=-1,
            width=self._width,
            height=self._height,
            title=self._title,
        )
        # windows peer must be created before tree control is added; otherwise,
        # the tree control will not work correctly. Some nodes seem not be added or visible.
        _ = Dialogs.create_dialog_peer(self._dialog)

        self._tree1 = Dialogs.insert_tree_control(
            dialog_ctrl=self._dialog,
            x=self._padding,
            y=self._padding,
            width=self._width - (self._padding * 2),
            height=self._height - (self._padding * 3),
            border=self._border_kind,
        )
        self._root1 = self._tree1.create_root(display_value="Root 1")
        if self._root1:
            _ = self._tree1.add_sub_node(parent_node=self._root1, display_value="Node 1")
            _ = self._tree1.add_sub_node(parent_node=self._root1, display_value="Node 3")
            _ = self._tree1.add_sub_node(parent_node=self._root1, display_value="Node 3")
        self._root2 = self._tree1.create_root(display_value="Root 2")
        if self._root2:
            _ = self._tree1.add_sub_node(parent_node=self._root2, display_value="Node 1")
            _ = self._tree1.add_sub_node(parent_node=self._root2, display_value="Node 3")
            _ = self._tree1.add_sub_node(parent_node=self._root2, display_value="Node 3")

        flat_list = [
            ["A1", "B1", "C1"],
            ["A1", "B1", "C2"],
            ["A1", "B2", "C3"],
            ["A2", "B3", "C4"],
            ["A2", "B3", "C5"],
            ["A2", "B3", "C6"],
        ]
        self._tree1.add_sub_tree(flat_tree=flat_list, parent_node=None)
        self._tree1.add_sub_tree(flat_tree=flat_list, parent_node=self._root2)
        # tree = self._tree1.convert_to_tree(flat_list)

        self._tree1.add_event_mouse_entered(self._fn_on_mouse_entered)
        self._tree1.add_event_mouse_exited(self._fn_on_mouse_exit)
        self._tree1.add_event_selection_changed(self._fn__on_tree_selection_changed)

    # region search
    def _search(self) -> None:
        se = SearchTree("C5", match_all=True)
        se.match_all = False
        se.register_rule(RuleDataInsensitive)
        se.register_rule(RuleTextInsensitive)
        # se.register_rule(RuleDataSensitive)
        result = se.find_node(self._root2)
        if result:
            print("Search Result:", result)
        else:
            print("Search Value, Not found")

    # endregion search

    # region Event Handlers
    def _init_handlers(self) -> None:
        def _on_mouse_entered(src: Any, event: EventArgs, control_src: Any, *args, **kwargs) -> None:
            self.on_mouse_entered(src, event, control_src, *args, **kwargs)

        def _on_mouse_exit(src: Any, event: EventArgs, control_src: Any, *args, **kwargs) -> None:
            self.on_mouse_exit(src, event, control_src, *args, **kwargs)

        def _on_tree_selection_changed(src: Any, event: EventArgs, control_src: Any, *args, **kwargs) -> None:
            self.on_tree_selection_changed(src, event, control_src, *args, **kwargs)

        self._fn_on_mouse_entered = _on_mouse_entered
        self._fn_on_mouse_exit = _on_mouse_exit
        self._fn__on_tree_selection_changed = _on_tree_selection_changed

    def on_mouse_entered(self, src: Any, event: EventArgs, control_src: Any, *args, **kwargs) -> None:
        # print(control_src)
        print("Mouse Entered:", control_src.name)

    def on_mouse_exit(self, src: Any, event: EventArgs, control_src: Any, *args, **kwargs) -> None:
        # print(control_src)
        print("Mouse Exited:", control_src.name)

    def on_tree_selection_changed(self, src: Any, event: EventArgs, control_src: Any, *args, **kwargs) -> None:
        # print(control_src)
        print("Selection changed:", control_src.name)

    # endregion Event Handlers

    def show(self) -> None:
        window = mLo.Lo.get_frame().getContainerWindow()
        ps = window.getPosSize()
        x = round(ps.Width / 2 - self._width / 2)
        y = round(ps.Height / 2 - self._height / 2)
        self._dialog.setTitle(self._title)
        self._dialog.setPosSize(x, y, self._width, self._height, PosSize.POSSIZE)
        self._dialog.setVisible(True)
        self._search()
        self._dialog.execute()
        self._dialog.dispose()


def main():
    with mLo.Lo.Loader(mLo.Lo.ConnectSocket(), opt=mLo.Lo.Options(verbose=True)):
        doc = Calc.create_doc()
        GUI.set_visible(visible=True, doc=doc)
        run()


def run() -> None:
    inst = Tree()
    inst.show()


if __name__ == "__main__":
    main()
