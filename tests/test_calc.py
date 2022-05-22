import pytest
from pathlib import Path

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])
# region    Sheet Methods
def test_get_sheet() -> None:
    # get_sheet is overload method.
    # testiing each overload.
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    # from ooodev.utils.gui import GUI
    with Lo.Loader() as loader:
        assert loader is not None
        doc = Calc.create_doc(loader)
        assert doc is not None
        # GUI.set_visible(is_visible=True, odoc=doc)
        sheet_names = Calc.get_sheet_names(doc)
        assert len(sheet_names) == 1
        assert sheet_names[0] == "Sheet1"
        # test overloads
        sheet_1_1 = Calc.get_sheet(doc=doc, sheet_name="Sheet1")
        assert sheet_1_1 is not None
        name_1_1 = Calc.get_sheet_name(sheet_1_1)
        assert name_1_1 == "Sheet1"

        sheet_1_2 = Calc.get_sheet(doc, "Sheet1")
        assert sheet_1_2 is not None
        name_1_2 = Calc.get_sheet_name(sheet_1_2)
        assert name_1_1 == name_1_2

        sheet_1_3 = Calc.get_sheet(doc=doc, index=0)
        assert sheet_1_3 is not None
        name_1_3 = Calc.get_sheet_name(sheet_1_3)
        assert name_1_3 == name_1_1

        sheet_1_4 = Calc.get_sheet(doc, 0)
        assert sheet_1_4 is not None
        name_1_4 = Calc.get_sheet_name(sheet_1_4)
        assert name_1_4 == name_1_1
        # Lo.delay(2000)
        Lo.close_doc(doc=doc, deliver_ownership=False)


def test_insert_sheet() -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    with Lo.Loader() as loader:
        assert loader is not None
        doc = Calc.create_doc(loader)
        assert doc is not None

        sheet_names = Calc.get_sheet_names(doc)
        assert len(sheet_names) == 1
        assert sheet_names[0] == "Sheet1"

        new_sht = Calc.insert_sheet(doc, "mysheet", 1)
        assert new_sht is not None

        sheet_names = Calc.get_sheet_names(doc)
        assert len(sheet_names) == 2
        assert sheet_names[1] == "mysheet"


def test_remove_sheet() -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    sheet_name = "mysheet"

    def add_sheet(doc):
        new_sht = Calc.insert_sheet(doc, sheet_name, 1)
        return new_sht

    with Lo.Loader() as loader:
        assert loader is not None
        doc = Calc.create_doc(loader)
        assert doc is not None

        sheet_names = Calc.get_sheet_names(doc)

        new_sht = add_sheet(doc)
        assert new_sht is not None

        sheet_names = Calc.get_sheet_names(doc)
        assert len(sheet_names) == 2
        assert sheet_names[1] == sheet_name
        # test overloads
        # test named args
        assert Calc.remove_sheet(doc=doc, sheet_name=sheet_name)

        new_sht = add_sheet(doc)
        assert new_sht is not None
        sheet_names = Calc.get_sheet_names(doc)
        assert len(sheet_names) == 2
        assert sheet_names[1] == sheet_name
        # test positional args
        assert Calc.remove_sheet(doc, sheet_name)

        new_sht = add_sheet(doc)
        assert new_sht is not None
        sheet_names = Calc.get_sheet_names(doc)
        assert len(sheet_names) == 2
        assert sheet_names[1] == sheet_name
        # test named args
        assert Calc.remove_sheet(doc=doc, index=1)

        new_sht = add_sheet(doc)
        assert new_sht is not None
        sheet_names = Calc.get_sheet_names(doc)
        assert len(sheet_names) == 2
        assert sheet_names[1] == sheet_name
        # test positional args
        assert Calc.remove_sheet(doc, 1)

        # sheet no longer exist
        assert Calc.remove_sheet(doc, 1) is False
        assert Calc.remove_sheet(doc, sheet_name) is False

        # incorrect number of arguments
        assert Calc.remove_sheet(doc) is False
        assert Calc.remove_sheet(doc, 0, "any") is False

        # incorrect type
        assert Calc.remove_sheet(doc, 22.3) is False


def test_move_sheet() -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    with Lo.Loader() as loader:
        assert loader is not None
        doc = Calc.create_doc(loader)
        assert doc is not None

        sheet_names = Calc.get_sheet_names(doc)
        assert len(sheet_names) == 1
        assert sheet_names[0] == "Sheet1"

        new_sht = Calc.insert_sheet(doc, "mysheet", 1)
        assert new_sht is not None

        sheet_names = Calc.get_sheet_names(doc)
        assert len(sheet_names) == 2
        assert sheet_names[1] == "mysheet"

        assert Calc.move_sheet(doc, "mysheet", 0)

        sheet_names = Calc.get_sheet_names(doc)
        assert len(sheet_names) == 2
        assert sheet_names[0] == "mysheet"
        assert sheet_names[1] == "Sheet1"


def test_set_sheet_name():
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    with Lo.Loader() as loader:
        assert loader is not None
        doc = Calc.create_doc(loader)
        assert doc is not None

        sheet_names = Calc.get_sheet_names(doc)
        assert len(sheet_names) == 1
        assert sheet_names[0] == "Sheet1"

        sheet = Calc.get_sheet(doc, 0)
        Calc.set_sheet_name(sheet, "mysheet")

        sheet_names = Calc.get_sheet_names(doc)
        assert len(sheet_names) == 1
        assert sheet_names[0] == "mysheet"


# endregion Sheet Methods

# region    View Methods
def test_get_controller() -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc
    from ooodev.utils.info import Info

    with Lo.Loader() as loader:
        assert loader is not None
        doc = Calc.create_doc(loader)
        assert doc is not None
        controller = Calc.get_controller(doc)
        assert controller is not None
        assert Info.is_type_interface(controller, "com.sun.star.frame.XController")


def test_zoom_value() -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc
    from ooodev.utils.gui import GUI

    with Lo.Loader() as loader:
        assert loader is not None
        doc = Calc.create_doc(loader)
        assert doc is not None
        GUI.set_visible(is_visible=True, odoc=doc)
        Lo.delay(500)
        Calc.zoom_value(doc, 160)
        Lo.delay(1000)


def test_zoom() -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc
    from ooodev.utils.gui import GUI

    with Lo.Loader() as loader:
        assert loader is not None
        doc = Calc.create_doc(loader)
        assert doc is not None
        Calc.zoom_value(doc, 250)
        GUI.set_visible(is_visible=True, odoc=doc)
        Lo.delay(500)
        Calc.zoom(doc, GUI.ZoomEnum.ENTIRE_PAGE)
        Lo.delay(1000)


def test_get_view() -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    with Lo.Loader() as loader:
        assert loader is not None
        doc = Calc.create_doc(loader)
        assert doc is not None

        view = Calc.get_view(doc)
        assert view is not None
        sheet = view.getActiveSheet()
        name = Calc.get_sheet_name(sheet)
        assert name == "Sheet1"


def test_get_set_active_sheet() -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    with Lo.Loader() as loader:
        assert loader is not None
        doc = Calc.create_doc(loader)
        assert doc is not None

        sheet = Calc.get_active_sheet(doc)
        name = Calc.get_sheet_name(sheet)
        assert name == "Sheet1"

        new_sht = Calc.insert_sheet(doc, "mysheet", 1)
        assert new_sht is not None

        lst_sht = Calc.insert_sheet(doc, "last", 1)
        assert lst_sht is not None

        Calc.set_active_sheet(doc, new_sht)
        asht = Calc.get_active_sheet(doc)
        name = Calc.get_sheet_name(asht)
        assert name == "mysheet"


def test_freeze() -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc
    from ooodev.utils.gui import GUI

    with Lo.Loader() as loader:
        assert loader is not None
        doc = Calc.create_doc(loader)
        assert doc is not None
        Calc.zoom(doc, GUI.ZoomEnum.ENTIRE_PAGE)
        GUI.set_visible(is_visible=True, odoc=doc)
        Calc.freeze(doc=doc, num_cols=2, num_rows=3)
        Lo.delay(1500)
        Lo.close_doc(doc=doc, deliver_ownership=False)


def test_freeze_cols() -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc
    from ooodev.utils.gui import GUI

    with Lo.Loader() as loader:
        assert loader is not None
        doc = Calc.create_doc(loader)
        assert doc is not None
        Calc.zoom_value(doc, 100)
        GUI.set_visible(is_visible=True, odoc=doc)
        Calc.freeze_cols(doc=doc, num_cols=2)
        Lo.delay(1500)
        Lo.close_doc(doc=doc, deliver_ownership=False)


def test_freeze_rows() -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc
    from ooodev.utils.gui import GUI

    with Lo.Loader() as loader:
        assert loader is not None
        doc = Calc.create_doc(loader)
        assert doc is not None
        Calc.zoom_value(doc, 100)
        GUI.set_visible(is_visible=True, odoc=doc)
        Calc.freeze_rows(doc=doc, num_rows=3)
        Lo.delay(1500)
        Lo.close_doc(doc=doc, deliver_ownership=False)


def test_goto_cell() -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    with Lo.Loader() as loader:
        assert loader is not None
        doc = Calc.create_doc(loader)
        assert doc is not None

        # test overloads
        Calc.goto_cell(cell_name="B4", doc=doc)
        cell = Calc.get_selected_cell_addr(doc=doc)
        assert cell.Column == 1
        assert cell.Row == 3

        Calc.goto_cell("A2", doc)
        cell = Calc.get_selected_cell_addr(doc=doc)
        assert cell.Column == 0
        assert cell.Row == 1

        frame = Calc.get_controller(doc).getFrame()
        Calc.goto_cell(cell_name="D5", frame=frame)
        cell = Calc.get_selected_cell_addr(doc=doc)
        assert cell.Column == 3
        assert cell.Row == 4

        Calc.goto_cell("E6", frame)
        cell = Calc.get_selected_cell_addr(doc=doc)
        assert cell.Column == 4
        assert cell.Row == 5


def test_split_window() -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc
    from ooodev.utils.gui import GUI

    with Lo.Loader() as loader:
        assert loader is not None
        doc = Calc.create_doc(loader)
        assert doc is not None
        Calc.zoom_value(doc, 100)
        GUI.set_visible(is_visible=True, odoc=doc)
        Calc.split_window(doc, "C4")
        Lo.delay(1500)


def test_get_selected_addr() -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc
    from com.sun.star.view import XSelectionSupplier
    from com.sun.star.frame import XModel

    with Lo.Loader() as loader:
        assert loader is not None
        doc = Calc.create_doc(loader)
        assert doc is not None

        sheet = Calc.get_active_sheet(doc)
        rng = Calc.get_cell_range(sheet=sheet, range_name="B1:D4")

        # rq = Lo.qi(XCellRangesQuery, sheet)
        # if rq:
        #     ranges = rq.queryVisibleCells()
        #     rng = ranges.getByIndex(0)

        cursor = sheet.createCursorByRange(rng)
        # cursor.collapseToSize(2, 4)

        controller = Calc.get_controller(doc)
        sel = Lo.qi(XSelectionSupplier, controller)
        if sel:
            sel.select(cursor)

        # test overloads
        addr = Calc.get_selected_addr(doc=doc)
        assert addr.StartColumn == 1
        assert addr.EndColumn == 3
        assert addr.StartRow == 0
        assert addr.EndRow == 3

        addr = Calc.get_selected_addr(doc)
        assert addr.StartColumn == 1
        assert addr.EndColumn == 3
        assert addr.StartRow == 0
        assert addr.EndRow == 3

        model = Lo.qi(XModel, doc)
        addr = Calc.get_selected_addr(model=model)
        assert addr.StartColumn == 1
        assert addr.EndColumn == 3
        assert addr.StartRow == 0
        assert addr.EndRow == 3

        addr = Calc.get_selected_addr(model)
        assert addr.StartColumn == 1
        assert addr.EndColumn == 3
        assert addr.StartRow == 0
        assert addr.EndRow == 3


def test_get_selected_cell_addr() -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc
    from com.sun.star.view import XSelectionSupplier

    with Lo.Loader() as loader:
        assert loader is not None
        doc = Calc.create_doc(loader)
        assert doc is not None

        sheet = Calc.get_active_sheet(doc)
        rng = Calc.get_cell_range(sheet=sheet, range_name="B1:D4")

        cursor = sheet.createCursorByRange(rng)

        controller = Calc.get_controller(doc)
        sel = Lo.qi(XSelectionSupplier, controller)
        if sel:
            sel.select(cursor)
        # should be None when more then a single cell is selected
        addr = Calc.get_selected_cell_addr(doc)
        assert addr is None

        Calc.goto_cell(cell_name="B4", doc=doc)
        addr = Calc.get_selected_cell_addr(doc)
        assert addr.Column == 1
        assert addr.Row == 3


# endregion View Methods

# region    view data methods
def test_get_view_panes() -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    with Lo.Loader() as loader:
        assert loader is not None
        doc = Calc.create_doc(loader)
        assert doc is not None
        panes = Calc.get_view_panes(doc)
        assert panes is not None
        pane = panes[0]
        assert pane.getFirstVisibleColumn() == 0
        assert pane.getFirstVisibleRow() == 0

def test_get_set_view_data() -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    with Lo.Loader() as loader:
        assert loader is not None
        doc = Calc.create_doc(loader)
        
        data = Calc.get_view_data(doc)
        assert data is not None
        # '100/60/0;0;tw:270;0/0/0/0/0/0/2/0/0/0/0'
        result = Calc.set_view_data(doc=doc, view_data=data)
        assert result is None

def test_get_set_view_states() -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    with Lo.Loader() as loader:
        assert loader is not None
        doc = Calc.create_doc(loader)
        
        data = Calc.get_view_states(doc)
        assert data is not None
        result = Calc.set_view_states(doc=doc, states=data)
        assert result is None
# endregion view data methods
