from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])
from ooodev.utils.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.calc import Calc
from ooodev.listeners.x_selection_change_adapter import XSelectionChangeAdapter

from com.sun.star.frame import XController
from com.sun.star.lang import EventObject
from com.sun.star.sheet import XSpreadsheet
from com.sun.star.sheet import XSpreadsheetDocument
from com.sun.star.table import CellAddress  # struct
from com.sun.star.view import XSelectionSupplier


class SelectionChangeListener(XSelectionChangeAdapter):
    def __init__(self, loader, visible: bool) -> None:

        self.loader = loader
        self.doc = Calc.create_doc(loader=self.loader)
        self.sheet = Calc.get_sheet(doc=self.doc, index=0)

        if visible:
            GUI.set_visible(is_visible=visible, odoc=self.doc)

        # insert some data
        Calc.set_col(sheet=self.sheet, cell_name="A1", values=("Smith", 42, 58.9, -66.5, 43.4, 44.5, 45.3))

        self.curr_addr = Calc.get_selected_cell_addr(self.doc)
        self.curr_val = self.get_cell_float(self.sheet, self.curr_addr)

        self.attach_listener(doc=self.doc)

    def attach_listener(self, doc: XSpreadsheetDocument) -> None:
        # start listening for selection changes
        ctrl = Calc.get_controller(doc)
        supp = Lo.qi(XSelectionSupplier, ctrl)
        supp.addSelectionChangeListener(self)

    def disposing(self, event: EventObject) -> None:
        print("Disposing")

    def selectionChanged(self, event: EventObject) -> None:
        ctrl = Lo.qi(XController, event.Source)
        if ctrl is None:
            print("No ctrl for event source")
            return
        addr = Calc.get_selected_cell_addr(self.doc)
        if addr is None:
            return
        if not Calc.is_equal_addresses(addr, self.curr_addr):
            self.curr_addr = addr
            self.curr_val = self.get_cell_float(sheet=self.sheet, addr=self.curr_addr)

    def get_cell_float(self, sheet: XSpreadsheet, addr: CellAddress) -> float | None:
        obj = Calc.get_val(sheet=sheet, addr=addr)
        if isinstance(obj, float):
            return obj
        return None


def test_modify_listener(loader) -> None:
    visible = False
    delay = 0  # 700
    try:
        sl = SelectionChangeListener(loader, visible)
        assert sl.curr_addr.Column == 0
        assert sl.curr_addr.Row == 0
        Lo.delay(delay)
        Calc.goto_cell(cell_name="A2", doc=sl.doc)
        assert sl.curr_addr.Column == 0
        assert sl.curr_addr.Row == 1
        assert sl.curr_val == 42.0
        Lo.delay(delay)
        Calc.goto_cell(cell_name="A3", doc=sl.doc)
        assert sl.curr_addr.Column == 0
        assert sl.curr_addr.Row == 2
        assert sl.curr_val == 58.9
        Lo.delay(delay)
        Calc.goto_cell(cell_name="A4", doc=sl.doc)
        assert sl.curr_addr.Column == 0
        assert sl.curr_addr.Row == 3
        assert sl.curr_val == -66.5
        Lo.delay(delay)
        Calc.goto_cell(cell_name="A5", doc=sl.doc)
        assert sl.curr_addr.Column == 0
        assert sl.curr_addr.Row == 4
        assert sl.curr_val == 43.4
        Lo.delay(delay)
        Calc.goto_cell(cell_name="A6", doc=sl.doc)
        assert sl.curr_addr.Column == 0
        assert sl.curr_addr.Row == 5
        assert sl.curr_val == 44.5
        Lo.delay(delay)
        Calc.goto_cell(cell_name="A7", doc=sl.doc)
        assert sl.curr_addr.Column == 0
        assert sl.curr_addr.Row == 6
        assert sl.curr_val == 45.3
        Lo.delay(delay)

    finally:
        Lo.close(closeable=sl.doc, deliver_ownership=False)
