import pytest
from pathlib import Path

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])
# region    Sheet Methods
def test_get_sheet(loader) -> None:
    # get_sheet is overload method.
    # testiing each overload.
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    # from ooodev.utils.gui import GUI

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


def test_insert_sheet(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

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
    Lo.close_doc(doc=doc, deliver_ownership=False)


def test_remove_sheet(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    sheet_name = "mysheet"

    def add_sheet(doc):
        new_sht = Calc.insert_sheet(doc, sheet_name, 1)
        return new_sht

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
    Lo.close_doc(doc=doc, deliver_ownership=False)


def test_move_sheet(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

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
    Lo.close_doc(doc=doc, deliver_ownership=False)


def test_set_sheet_name(loader):
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

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
    Lo.close_doc(doc=doc, deliver_ownership=False)


# endregion Sheet Methods

# region    View Methods
def test_get_controller(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc
    from ooodev.utils.info import Info

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    controller = Calc.get_controller(doc)
    assert controller is not None
    assert Info.is_type_interface(controller, "com.sun.star.frame.XController")
    Lo.close_doc(doc=doc, deliver_ownership=False)


def test_zoom_value(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc
    from ooodev.utils.gui import GUI

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    GUI.set_visible(is_visible=True, odoc=doc)
    Lo.delay(500)
    Calc.zoom_value(doc, 160)
    Lo.delay(1000)
    Lo.close_doc(doc=doc, deliver_ownership=False)


def test_zoom(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc
    from ooodev.utils.gui import GUI

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    Calc.zoom_value(doc, 250)
    GUI.set_visible(is_visible=True, odoc=doc)
    Lo.delay(500)
    Calc.zoom(doc, GUI.ZoomEnum.ENTIRE_PAGE)
    Lo.delay(1000)
    Lo.close_doc(doc=doc, deliver_ownership=False)


def test_get_view(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None

    view = Calc.get_view(doc)
    assert view is not None
    sheet = view.getActiveSheet()
    name = Calc.get_sheet_name(sheet)
    assert name == "Sheet1"
    Lo.close_doc(doc=doc, deliver_ownership=False)


def test_get_set_active_sheet(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

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
    Lo.close_doc(doc=doc, deliver_ownership=False)


def test_freeze(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc
    from ooodev.utils.gui import GUI

    # with Lo.Loader() as loader:
    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    Calc.zoom(doc, GUI.ZoomEnum.ENTIRE_PAGE)
    GUI.set_visible(is_visible=True, odoc=doc)
    Calc.freeze(doc=doc, num_cols=2, num_rows=3)
    Lo.delay(1500)
    Lo.close_doc(doc=doc, deliver_ownership=False)


def test_freeze_cols(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc
    from ooodev.utils.gui import GUI

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    Calc.zoom_value(doc, 100)
    GUI.set_visible(is_visible=True, odoc=doc)
    Calc.freeze_cols(doc=doc, num_cols=2)
    Lo.delay(1500)
    Lo.close_doc(doc=doc, deliver_ownership=False)


def test_freeze_rows(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc
    from ooodev.utils.gui import GUI

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    Calc.zoom_value(doc, 100)
    GUI.set_visible(is_visible=True, odoc=doc)
    Calc.freeze_rows(doc=doc, num_rows=3)
    Lo.delay(1500)
    Lo.close_doc(doc=doc, deliver_ownership=False)


def test_goto_cell(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

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
    Lo.close_doc(doc=doc, deliver_ownership=False)


def test_split_window(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc
    from ooodev.utils.gui import GUI

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    Calc.zoom_value(doc, 100)
    GUI.set_visible(is_visible=True, odoc=doc)
    Calc.split_window(doc, "C4")
    Lo.delay(1500)
    Lo.close_doc(doc=doc, deliver_ownership=False)


def test_get_selected_addr(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc
    from com.sun.star.view import XSelectionSupplier
    from com.sun.star.frame import XModel

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
    Lo.close_doc(doc=doc, deliver_ownership=False)


def test_get_selected_cell_addr(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc
    from com.sun.star.view import XSelectionSupplier

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
    Lo.close_doc(doc=doc, deliver_ownership=False)


# endregion View Methods

# region    view data methods
def test_get_view_panes(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    panes = Calc.get_view_panes(doc)
    assert panes is not None
    pane = panes[0]
    assert pane.getFirstVisibleColumn() == 0
    assert pane.getFirstVisibleRow() == 0
    Lo.close_doc(doc=doc, deliver_ownership=False)


def test_get_set_view_data(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    assert loader is not None
    doc = Calc.create_doc(loader)

    data = Calc.get_view_data(doc)
    assert data is not None
    # '100/60/0;0;tw:270;0/0/0/0/0/0/2/0/0/0/0'
    result = Calc.set_view_data(doc=doc, view_data=data)
    assert result is None
    Lo.close_doc(doc=doc, deliver_ownership=False)


def test_get_set_view_states(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    assert loader is not None
    doc = Calc.create_doc(loader)

    data = Calc.get_view_states(doc)
    assert data is not None
    result = Calc.set_view_states(doc=doc, states=data)
    assert result is None
    Lo.close_doc(doc=doc, deliver_ownership=False)


# endregion view data methods


# region insert/remove rows, columns, cells
def test_insert_row(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    # from ooodev.utils.gui import GUI

    assert loader is not None
    doc = Calc.create_doc(loader)
    sheet = Calc.get_active_sheet(doc)
    Calc.set_val(value="hello world", sheet=sheet, cell_name="A2")
    # GUI.set_visible(is_visible=True, odoc=doc)
    Calc.insert_row(sheet=sheet, idx=1)
    val = Calc.get_val(sheet=sheet, cell_name="A3")
    # Lo.delay(1500)
    assert val == "hello world"
    Lo.close(closeable=doc, deliver_ownership=False)


def test_delete_row(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    # from ooodev.utils.gui import GUI

    assert loader is not None
    doc = Calc.create_doc(loader)
    sheet = Calc.get_active_sheet(doc)
    Calc.set_val(value="hello world", sheet=sheet, cell_name="A3")
    # GUI.set_visible(is_visible=True, odoc=doc)
    Calc.delete_row(sheet=sheet, idx=1)
    val = Calc.get_val(sheet=sheet, cell_name="A2")
    # Lo.delay(1500)
    assert val == "hello world"
    Lo.close(closeable=doc, deliver_ownership=False)


def test_insert_column(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    # from ooodev.utils.gui import GUI

    assert loader is not None
    doc = Calc.create_doc(loader)
    sheet = Calc.get_active_sheet(doc)
    Calc.set_val(value="hello world", sheet=sheet, cell_name="C2")
    # GUI.set_visible(is_visible=True, odoc=doc)
    Calc.delete_column(sheet=sheet, idx=1)
    val = Calc.get_val(sheet=sheet, cell_name="B2")
    # Lo.delay(1500)
    Lo.close(closeable=doc, deliver_ownership=False)
    assert val == "hello world"


def test_insert_cells_down(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    # from ooodev.utils.gui import GUI

    assert loader is not None
    doc = Calc.create_doc(loader)
    sheet = Calc.get_active_sheet(doc)

    # GUI.set_visible(is_visible=True, odoc=doc)
    rng = Calc.get_cell_range(sheet=sheet, range_name="B1:D4")
    Calc.set_val(value="hello world", sheet=sheet, cell_name="B5")
    Calc.insert_cells(sheet=sheet, cell_range=rng, is_shift_right=False)
    # Lo.delay(1500)
    val = Calc.get_val(sheet=sheet, cell_name="B9")
    Lo.close(closeable=doc, deliver_ownership=False)
    assert val == "hello world"


def test_insert_cells_right(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    # from ooodev.utils.gui import GUI

    assert loader is not None
    doc = Calc.create_doc(loader)
    sheet = Calc.get_active_sheet(doc)

    # GUI.set_visible(is_visible=True, odoc=doc)
    rng = Calc.get_cell_range(sheet=sheet, range_name="B1:D4")
    Calc.set_val(value="hello world", sheet=sheet, cell_name="D4")
    Calc.insert_cells(sheet=sheet, cell_range=rng, is_shift_right=True)
    # Lo.delay(1500)
    val = Calc.get_val(sheet=sheet, cell_name="G4")
    Lo.close(closeable=doc, deliver_ownership=False)
    assert val == "hello world"


def test_delete_cells_down(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    # from ooodev.utils.gui import GUI

    assert loader is not None
    doc = Calc.create_doc(loader)
    sheet = Calc.get_active_sheet(doc)

    # GUI.set_visible(is_visible=True, odoc=doc)
    rng = Calc.get_cell_range(sheet=sheet, range_name="B1:D4")
    Calc.set_val(value="hello world", sheet=sheet, cell_name="B9")
    # Calc.set_val(value="A1",sheet=sheet, cell_name="A1")
    Calc.delete_cells(sheet=sheet, cell_range=rng, is_shift_left=False)
    # Lo.delay(1500)
    # note without calling 'sheet = Calc.get_active_sheet(doc)' this test fails.
    # for unknown reason Calc.get_val() will always return None after delete_cell()
    # refreshing sheet solves the issue.
    sheet = Calc.get_active_sheet(doc)
    val = Calc.get_val(sheet=sheet, cell_name="B5")
    Lo.close(closeable=doc, deliver_ownership=False)
    assert val == "hello world"


def test_delete_cells_left(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    # from ooodev.utils.gui import GUI

    assert loader is not None
    doc = Calc.create_doc(loader)
    sheet = Calc.get_active_sheet(doc)

    # GUI.set_visible(is_visible=True, odoc=doc)
    rng = Calc.get_cell_range(sheet=sheet, range_name="B1:D4")
    Calc.set_val(value="hello world", sheet=sheet, cell_name="B5")
    # Calc.set_val(value="A1",sheet=sheet, cell_name="A1")
    Calc.delete_cells(sheet=sheet, cell_range=rng, is_shift_left=False)
    # Lo.delay(1500)
    # note without calling 'sheet = Calc.get_active_sheet(doc)' this test fails.
    # for unknown reason Calc.get_val() will always return None after delete_cell()
    # refreshing sheet solves the issue.
    sheet = Calc.get_active_sheet(doc)
    val = Calc.get_val(sheet=sheet, cell_name="B1")
    Lo.close(closeable=doc, deliver_ownership=False)
    assert val == "hello world"


# endregion insert/remove rows, columns, cells

# region set/get values in cells
def test_set_val(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    # from ooodev.utils.gui import GUI

    assert loader is not None
    doc = Calc.create_doc(loader)
    sheet = Calc.get_active_sheet(doc)

    # test overloads
    # set_val(value: object, sheet: XSpreadsheet, cell_name: str)
    Calc.set_val(value="one", sheet=sheet, cell_name="A2")
    val = Calc.get_val(sheet=sheet, cell_name="A2")
    assert val == "one"
    Calc.set_val("one", sheet, "B2")
    val = Calc.get_val(sheet, "B2")
    assert val == "one"

    # set_val(value: object, sheet: XSpreadsheet, column: int, row: int)
    Calc.set_val(value="two", sheet=sheet, column=1, row=2)
    val = Calc.get_val(sheet=sheet, column=1, row=2)
    assert val == "two"
    Calc.set_val("two", sheet, 1, 3)
    val = Calc.get_val(sheet, 1, 3)
    assert val == "two"

    # set_val(value: object, cell: XCell)
    cell = Calc.get_cell(sheet=sheet, column=2, row=4)
    Calc.set_val(value="three", cell=cell)
    val = Calc.get_val(sheet=sheet, column=2, row=4)
    assert val == "three"
    cell = Calc.get_cell(sheet, 2, 5)
    Calc.set_val("three", cell)
    val = Calc.get_val(sheet, 2, 5)
    assert val == "three"
    Lo.close(closeable=doc, deliver_ownership=False)

def test_get_val(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    # from ooodev.utils.gui import GUI

    assert loader is not None
    doc = Calc.create_doc(loader)
    sheet = Calc.get_active_sheet(doc)

    # test overloads
    Calc.set_val(value="one", sheet=sheet, cell_name="A2")
    
    # test overlaods
    # get_val(sheet: XSpreadsheet, cell_name: str)
    val = Calc.get_val(sheet=sheet, cell_name="A2")
    assert val == "one"
    val = Calc.get_val(sheet, "A2")
    assert val == "one"
    
    # get_val(sheet: XSpreadsheet, addr: CellAddress)
    addr = Calc.get_cell_address(sheet=sheet, cell_name="A2")
    val = Calc.get_val(sheet=sheet, addr=addr)
    assert val == "one"
    addr = Calc.get_cell_address(sheet, "A2")
    val = Calc.get_val(sheet, addr)
    assert val == "one"
    
    # get_val(sheet: XSpreadsheet, column: int, row: int)
    val = Calc.get_val(sheet=sheet, column=0, row=1)
    assert val == "one"
    val = Calc.get_val(sheet, 0, 1)
    assert val == "one"
    
    # def get_val(cell: XCell)
    cell = Calc.get_cell(sheet=sheet, column=0, row=1)
    val = Calc.get_val(cell=cell)
    assert val == "one"
    cell = Calc.get_cell(sheet, 0, 1)
    val = Calc.get_val(cell)
    assert val == "one"
    
    Lo.close(closeable=doc, deliver_ownership=False)

def test_get_num(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc
    # from ooodev.utils.gui import GUI

    assert loader is not None
    doc = Calc.create_doc(loader)
    sheet = Calc.get_active_sheet(doc)
    Calc.set_val(value=1, sheet=sheet, cell_name="A1")
    Calc.set_val(value=33.87, sheet=sheet, cell_name="A2")
    Calc.set_val(value="Not a Number", sheet=sheet, cell_name="A3")
    # test overloads
    # def get_num(cell: XCell)
    cell = Calc.get_cell(sheet=sheet, column=0, row=0)
    val = Calc.get_num(cell=cell)
    assert val == 1.0
    cell = Calc.get_cell(sheet, 0, 1)
    val = Calc.get_num(cell)
    assert val == 33.87
    cell = Calc.get_cell(sheet, 0, 2)
    val = Calc.get_num(cell)
    assert val == 0.0
    cell = Calc.get_cell(sheet, 3, 3) # empty cell
    val = Calc.get_num(cell)
    assert val == 0.0

    # get_num(sheet: XSpreadsheet, cell_name: str)
    val = Calc.get_num(sheet=sheet, cell_name='A1')
    assert val == 1.0
    val = Calc.get_num(sheet,'A2')
    assert val == 33.87
    val = Calc.get_num(sheet,'A3')
    assert val == 0.0
    val = Calc.get_num(sheet,'C3') # empty cell
    assert val == 0.0

    # def get_num(sheet: XSpreadsheet, addr: CellAddress)
    addr = Calc.get_cell_address(sheet=sheet, cell_name="A1")
    val = Calc.get_num(sheet=sheet, addr=addr)
    assert val == 1.0
    addr = Calc.get_cell_address(sheet, "A2")
    val = Calc.get_num(sheet, addr)
    assert val == 33.87
    addr = Calc.get_cell_address(sheet, "A3")
    val = Calc.get_num(sheet, addr)
    assert val == 0.0
    addr = Calc.get_cell_address(sheet, "B3") # empty cell
    val = Calc.get_num(sheet, addr)
    assert val == 0.0
    
    # get_num(sheet: XSpreadsheet, column: int, row: int)
    val = Calc.get_num(sheet=sheet, column=0, row=0)
    assert val == 1.0
    val = Calc.get_num(sheet, 0, 1)
    assert val == 33.87
    val = Calc.get_num(sheet, 0, 2)
    assert val == 0.0
    val = Calc.get_num(sheet, 2, 2) # empty cell
    assert val == 0.0

    Lo.close(closeable=doc, deliver_ownership=False)


def test_get_str(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc
    # from ooodev.utils.gui import GUI

    assert loader is not None
    doc = Calc.create_doc(loader)
    sheet = Calc.get_active_sheet(doc)
    Calc.set_val(value="one", sheet=sheet, cell_name="A1")
    Calc.set_val(value=33.87, sheet=sheet, cell_name="A2")
    Calc.set_val(value="custom val", sheet=sheet, cell_name="A3")
    # test overloads
    # def get_string(cell: XCell)
    cell = Calc.get_cell(sheet=sheet, column=0, row=0)
    val = Calc.get_string(cell=cell)
    assert val == "one"
    cell = Calc.get_cell(sheet, 0, 1)
    val = Calc.get_string(cell)
    assert val == "33.87"
    cell = Calc.get_cell(sheet, 0, 2)
    val = Calc.get_string(cell)
    assert val == "custom val"
    cell = Calc.get_cell(sheet, 2, 2) # empty cell
    val = Calc.get_string(cell)
    assert val == ""

    # get_string(sheet: XSpreadsheet, cell_name: str)
    val = Calc.get_string(sheet=sheet, cell_name='A1')
    assert val == "one"
    val = Calc.get_string(sheet,'A2')
    assert val == "33.87"
    val = Calc.get_string(sheet,'A3')
    assert val == "custom val"
    val = Calc.get_string(sheet,'C3') # empty cell
    assert val == ""

    # def get_string(sheet: XSpreadsheet, addr: CellAddress)
    addr = Calc.get_cell_address(sheet=sheet, cell_name="A1")
    val = Calc.get_string(sheet=sheet, addr=addr)
    assert val == "one"
    addr = Calc.get_cell_address(sheet, "A2")
    val = Calc.get_string(sheet, addr)
    assert val == "33.87"
    addr = Calc.get_cell_address(sheet, "A3")
    val = Calc.get_string(sheet, addr)
    assert val == "custom val"
    addr = Calc.get_cell_address(sheet, "C3") # empty cell
    val = Calc.get_string(sheet, addr)
    assert val == ""
    
    # get_string(sheet: XSpreadsheet, column: int, row: int)
    val = Calc.get_string(sheet=sheet, column=0, row=0)
    assert val == "one"
    val = Calc.get_string(sheet, 0, 1)
    assert val == "33.87"
    val = Calc.get_string(sheet, 0, 2)
    assert val == "custom val"
    val = Calc.get_string(sheet, 2, 2) # empty cell
    assert val == ""

    Lo.close(closeable=doc, deliver_ownership=False)
# endregion set/get values in cells

# region set/get values in 2D array
def test_set_array_by_range(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc
    from ooodev.utils.gen_util import TableHelper
    # from ooodev.utils.gui import GUI

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    arr1 = (
        (1, 2, 3),
        (4, 5, 6),
        (7, 8 , 9)
    )
    arr = TableHelper.to_2d_list(arr1)
    sheet = Calc.get_active_sheet(doc=doc)
    Calc.set_array(sheet=sheet, name="A1:C3", values=arr)
    # GUI.set_visible(is_visible=True, odoc=doc)
    # Lo.delay(3500)
    val = Calc.get_num(sheet,'A1')
    assert val == 1.0
    val = Calc.get_num(sheet,'c2')
    assert val == 6.0
    val = Calc.get_num(sheet,'C3')
    assert val == 9.0
    
    arr_size = 50
    arr = TableHelper.to_2d_tuple(TableHelper.make_2d_array(arr_size, arr_size))
    rng = TableHelper.make_column_name(arr_size)
    Calc.set_array(sheet=sheet, name=f"A1:{rng}{arr_size}", values=arr)
    # Lo.delay(3500)
    val = Calc.get_num(sheet, f"{rng}{arr_size}")
    assert val == 1.0
    
    Lo.close(closeable=doc, deliver_ownership=False)


def test_set_array_by_cell(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc
    # from ooodev.utils.gui import GUI
    from ooodev.utils.gen_util import TableHelper
    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    arr = TableHelper.to_2d_tuple(TableHelper.to_tuple(2))
    sheet = Calc.get_active_sheet(doc=doc)
    Calc.set_array(sheet=sheet, name="B2", values=arr)
    # GUI.set_visible(is_visible=True, odoc=doc)
    # Lo.delay(1500)
    val = Calc.get_num(sheet,'B2')
    assert val == 2.0
    
    arr_size = 12
    arr = TableHelper.to_2d_tuple(TableHelper.make_2d_array(arr_size, arr_size, 45.7))
    rng = TableHelper.make_column_name(arr_size)
    Calc.set_array(sheet=sheet, name=f"A1:{rng}{arr_size}", values=arr)
    val = Calc.get_num(sheet, f"{rng}{arr_size}")
    assert val == 45.7
    
    # set as single cell
    arr_size = 12
    arr = TableHelper.to_2d_tuple(TableHelper.make_2d_array(arr_size, arr_size, 3.14))
    rng = TableHelper.make_column_name(arr_size)
    Calc.set_array(sheet=sheet, name=f"A1", values=arr)
    val = Calc.get_num(sheet, f"{rng}{arr_size}")
    assert val == 3.14
    
    Lo.close(closeable=doc, deliver_ownership=False)

def test_get_array(loader) -> None:
    def arr_cb(row:int, col:int, prev_val) -> float:
        if row == 0 and col == 0:
            return 1.0
        return prev_val + 1.0
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc
    # from ooodev.utils.gui import GUI
    from ooodev.utils.gen_util import TableHelper
    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    sheet = Calc.get_active_sheet(doc=doc)
    
    arr_size = 8
    arr = TableHelper.to_2d_tuple(TableHelper.make_2d_array(arr_size, arr_size,arr_cb))
    rng = TableHelper.make_column_name(arr_size)
    Calc.set_array(sheet=sheet, name=f"A1", values=arr)
    val = Calc.get_num(sheet, f"{rng}{arr_size}")
    assert val == float(arr_size * arr_size)
    
    rng_name = Calc.get_range_str(start_col=0, start_row=0, end_col= arr_size -1, end_row=arr_size -1)
    result_arr = Calc.get_array(sheet=sheet, range_name=rng_name)
    for row, row_data in enumerate(arr):
        for col, _ in enumerate(row_data):
            assert result_arr[row][col] == arr[row][col]
    
    Lo.close(closeable=doc, deliver_ownership=False)
# endregion set/get values in 2D array
