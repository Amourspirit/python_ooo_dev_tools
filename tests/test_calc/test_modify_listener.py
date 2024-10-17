from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])
import types
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.calc import Calc
from ooodev.listeners.x_top_window_adapter import XTopWindowAdapter
from ooodev.listeners.x_modify_adapter import XModifyAdapter

from com.sun.star.awt import XExtendedToolkit
from com.sun.star.lang import EventObject
from com.sun.star.sheet import XSpreadsheetDocument
from com.sun.star.util import XModifyBroadcaster


class ModifyListener(XModifyAdapter):
    def __init__(self, loader) -> None:
        visible = False
        delay = 0  # 1000
        self.loader = loader
        self.doc = Calc.create_doc(loader=self.loader)

        # region Top Window Listener
        # windowOpened only fires if visible is true
        def windowOpened(self, event: EventObject) -> None:
            print("Window Opened")

        # currently windowClosing does not fire in test.
        # most likely because of how test closes office
        def windowClosing(self, event: EventObject):
            print("Closing")

        tk = Lo.create_instance_mcf(XExtendedToolkit, "com.sun.star.awt.Toolkit")
        top_adapter = XTopWindowAdapter()
        top_adapter.windowClosing = types.MethodType(windowClosing, top_adapter)
        top_adapter.windowOpened = types.MethodType(windowOpened, top_adapter)
        tk.addTopWindowListener(top_adapter)
        # endregion Top Window Listener

        mb = Lo.qi(XModifyBroadcaster, self.doc)
        mb.addModifyListener(self)

        self.sheet = Calc.get_sheet(doc=self.doc, index=0)

        if visible:
            GUI.set_visible(is_visible=visible, odoc=self.doc)

        # insert some data
        Calc.set_col(sheet=self.sheet, cell_name="A1", values=("Smith", 42, 58.9, -66.5, 43.4, 44.5, 45.3))

        Lo.delay(delay)

    def disposing(self, event: EventObject) -> None:
        print("Disposing")

    def modified(self, event: EventObject) -> None:
        print("Modified")
        doc = Lo.qi(XSpreadsheetDocument, event.Source)
        sheet = Calc.get_active_sheet(doc)
        addr = Calc.get_selected_cell_addr(doc=doc)
        print(f"{Calc.get_cell_str(addr)} = {Calc.get_val(sheet, addr)}")


def test_modify_listener(loader, capsys: pytest.CaptureFixture) -> None:
    ml = ModifyListener(loader)
    capsys.readouterr()  # clear buffer
    Calc.goto_cell(cell_name="A8", doc=ml.doc)
    Calc.set_val(21, sheet=ml.sheet, cell_name="A8")
    capture = capsys.readouterr()
    presult = capture.out
    Lo.close(closeable=ml.doc, deliver_ownership=False)
    assert presult == "Modified\nA8 = 21.0\n"
