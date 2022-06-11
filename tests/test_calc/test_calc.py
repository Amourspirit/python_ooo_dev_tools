from typing import cast
import pytest

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
    with pytest.raises(TypeError):
        # unused keyword
        Calc.get_sheet(doc=doc, name="Sheet1")
    with pytest.raises(TypeError):
        # Incorrect number of params
        Calc.get_sheet(doc=doc)

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

    # incorrect type
    assert Calc.remove_sheet(doc, 22.3) is False
    with pytest.raises(TypeError):
        # unused keyword
        Calc.remove_sheet(doc=doc, other=1)
    with pytest.raises(TypeError):
        # incorrect number of arguments
        Calc.remove_sheet(doc)
    with pytest.raises(TypeError):
        # incorrect number of arguments
        Calc.remove_sheet(doc, 0, "any")

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

    with pytest.raises(TypeError):
        # incorrect number of params
        Calc.goto_cell(cell_name="D5")
    with pytest.raises(TypeError):
        # unused keyword
        Calc.goto_cell(cell_name="D5", f=frame)
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
    with pytest.raises(TypeError):
        # incorrect number of args
        Calc.get_selected_addr()
    with pytest.raises(TypeError):
        # unused keyword
        Calc.get_selected_addr(custom=1)
    Lo.close_doc(doc=doc, deliver_ownership=False)


def test_get_selected_cell_addr(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.exceptions.ex import CellError
    from ooodev.office.calc import Calc
    from com.sun.star.view import XSelectionSupplier

    assert loader is not None
    doc = Calc.create_doc(loader)

    sheet = Calc.get_active_sheet(doc)
    rng = Calc.get_cell_range(sheet=sheet, range_name="B1:D4")

    cursor = sheet.createCursorByRange(rng)

    controller = Calc.get_controller(doc)
    sel = Lo.qi(XSelectionSupplier, controller)
    if sel:
        sel.select(cursor)
    # should raise error when more then a single cell is selected
    with pytest.raises(CellError):
        addr = Calc.get_selected_cell_addr(doc)

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

# region    insert/remove rows, columns, cells
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

# region    set/get values in cells
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
    Calc.set_val(value="two", sheet=sheet, col=1, row=2)
    val = Calc.get_val(sheet=sheet, col=1, row=2)
    assert val == "two"
    Calc.set_val("two", sheet, 1, 3)
    val = Calc.get_val(sheet, 1, 3)
    assert val == "two"

    # set_val(value: object, cell: XCell)
    cell = Calc.get_cell(sheet=sheet, col=2, row=4)
    Calc.set_val(value="three", cell=cell)
    val = Calc.get_val(sheet=sheet, col=2, row=4)
    assert val == "three"
    cell = Calc.get_cell(sheet, 2, 5)
    Calc.set_val("three", cell)
    val = Calc.get_val(sheet, 2, 5)
    assert val == "three"
    with pytest.raises(TypeError):
        # error on unused key
        Calc.set_val(value="one", sheet=sheet, cell_name="A2", other=1)
    with pytest.raises(TypeError):
        # error on incorrect number of args
        Calc.set_val(sheet)
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
    val = Calc.get_val(sheet=sheet, col=0, row=1)
    assert val == "one"
    val = Calc.get_val(sheet, 0, 1)
    assert val == "one"

    # def get_val(cell: XCell)
    cell = Calc.get_cell(sheet=sheet, col=0, row=1)
    val = Calc.get_val(cell=cell)
    assert val == "one"
    cell = Calc.get_cell(sheet, 0, 1)
    val = Calc.get_val(cell)
    assert val == "one"

    with pytest.raises(TypeError):
        # error on unused key
        Calc.get_val(sheet=sheet, col=0, rew=1)
    with pytest.raises(TypeError):
        # error on incorrect number of args
        Calc.get_val(sheet, 0, 1, 3)

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
    cell = Calc.get_cell(sheet=sheet, col=0, row=0)
    val = Calc.get_num(cell=cell)
    assert val == 1.0
    cell = Calc.get_cell(sheet, 0, 1)
    val = Calc.get_num(cell)
    assert val == 33.87
    cell = Calc.get_cell(sheet, 0, 2)
    val = Calc.get_num(cell)
    assert val == 0.0
    cell = Calc.get_cell(sheet, 3, 3)  # empty cell
    val = Calc.get_num(cell)
    assert val == 0.0

    # get_num(sheet: XSpreadsheet, cell_name: str)
    val = Calc.get_num(sheet=sheet, cell_name="A1")
    assert val == 1.0
    val = Calc.get_num(sheet, "A2")
    assert val == 33.87
    val = Calc.get_num(sheet, "A3")
    assert val == 0.0
    val = Calc.get_num(sheet, "C3")  # empty cell
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
    addr = Calc.get_cell_address(sheet, "B3")  # empty cell
    val = Calc.get_num(sheet, addr)
    assert val == 0.0

    # get_num(sheet: XSpreadsheet, column: int, row: int)
    val = Calc.get_num(sheet=sheet, col=0, row=0)
    assert val == 1.0
    val = Calc.get_num(sheet, 0, 1)
    assert val == 33.87
    val = Calc.get_num(sheet, 0, 2)
    assert val == 0.0
    val = Calc.get_num(sheet, 2, 2)  # empty cell
    assert val == 0.0

    with pytest.raises(TypeError):
        # error on unused key
        Calc.get_num(sheet=sheet, col=0, rew=0)
    with pytest.raises(TypeError):
        # error on incorrect number of args
        Calc.get_num(sheet, 0, 1, 4)

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
    cell = Calc.get_cell(sheet=sheet, col=0, row=0)
    val = Calc.get_string(cell=cell)
    assert val == "one"
    cell = Calc.get_cell(sheet, 0, 1)
    val = Calc.get_string(cell)
    assert val == "33.87"
    cell = Calc.get_cell(sheet, 0, 2)
    val = Calc.get_string(cell)
    assert val == "custom val"
    cell = Calc.get_cell(sheet, 2, 2)  # empty cell
    val = Calc.get_string(cell)
    assert val == ""

    # get_string(sheet: XSpreadsheet, cell_name: str)
    val = Calc.get_string(sheet=sheet, cell_name="A1")
    assert val == "one"
    val = Calc.get_string(sheet, "A2")
    assert val == "33.87"
    val = Calc.get_string(sheet, "A3")
    assert val == "custom val"
    val = Calc.get_string(sheet, "C3")  # empty cell
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
    addr = Calc.get_cell_address(sheet, "C3")  # empty cell
    val = Calc.get_string(sheet, addr)
    assert val == ""

    # get_string(sheet: XSpreadsheet, column: int, row: int)
    val = Calc.get_string(sheet=sheet, col=0, row=0)
    assert val == "one"
    val = Calc.get_string(sheet, 0, 1)
    assert val == "33.87"
    val = Calc.get_string(sheet, 0, 2)
    assert val == "custom val"
    val = Calc.get_string(sheet, 2, 2)  # empty cell
    assert val == ""

    with pytest.raises(TypeError):
        # error on unused key
        Calc.get_string(sheet=sheet, cellName="A1")
    with pytest.raises(TypeError):
        # error on incorrect number of args
        Calc.get_string(sheet, 0, 1, 5)

    Lo.close(closeable=doc, deliver_ownership=False)


# endregion set/get values in cells

# region    set/get values in 2D array
def test_set_array_by_range(loader) -> None:
    # set_array(values: Sequence[Sequence[object]], sheet: XSpreadsheet, name: str)
    # test when name is a range
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc
    from ooodev.utils.gen_util import TableHelper

    # from ooodev.utils.gui import GUI

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    arr1 = ((1, 2, 3), (4, 5, 6), (7, 8, 9))
    arr = TableHelper.to_2d_list(arr1)
    sheet = Calc.get_active_sheet(doc=doc)
    Calc.set_array(sheet=sheet, name="A1:C3", values=arr)
    # GUI.set_visible(is_visible=True, odoc=doc)
    # Lo.delay(3500)
    val = Calc.get_num(sheet, "A1")
    assert val == 1.0
    val = Calc.get_num(sheet, "c2")
    assert val == 6.0
    val = Calc.get_num(sheet, "C3")
    assert val == 9.0

    arr_size = 25
    arr = TableHelper.to_2d_tuple(TableHelper.make_2d_array(arr_size, arr_size))
    rng = TableHelper.make_column_name(arr_size)
    Calc.set_array(sheet=sheet, name=f"A1:{rng}{arr_size}", values=arr)
    # Lo.delay(3500)
    val = Calc.get_num(sheet, f"{rng}{arr_size}")
    assert val == 1.0

    # positiona args
    Calc.set_array(arr, sheet, f"A1:{rng}{arr_size}")
    val = Calc.get_num(sheet, f"{rng}{arr_size}")
    assert val == 1.0

    Lo.close(closeable=doc, deliver_ownership=False)


def test_set_array_by_cell(loader) -> None:
    # set_array(values: Sequence[Sequence[object]], sheet: XSpreadsheet, name: str)
    # test when name is a cell
    def arr_cb(row: int, col: int, prev_val) -> float:
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
    arr = TableHelper.to_2d_tuple(TableHelper.to_tuple(2))
    sheet = Calc.get_active_sheet(doc=doc)
    Calc.set_array(sheet=sheet, name="B2", values=arr)
    # GUI.set_visible(is_visible=True, odoc=doc)
    # Lo.delay(1500)
    val = Calc.get_num(sheet, "B2")
    assert val == 2.0

    arr_size = 12
    arr = TableHelper.to_2d_tuple(TableHelper.make_2d_array(arr_size, arr_size, arr_cb))
    rng = TableHelper.make_column_name(arr_size)
    # keyword args
    Calc.set_array(sheet=sheet, name=f"A1", values=arr)
    val = Calc.get_num(sheet, f"{rng}{arr_size}")
    assert val == float(arr_size * arr_size)

    # positiona args
    Calc.set_array(arr, sheet, f"A1")
    val = Calc.get_num(sheet, f"{rng}{arr_size}")
    assert val == float(arr_size * arr_size)

    # set as single cell
    arr_size = 12
    arr = TableHelper.to_2d_tuple(TableHelper.make_2d_array(arr_size, arr_size, 3.14))
    rng = TableHelper.make_column_name(arr_size)
    Calc.set_array(sheet=sheet, name=f"A1", values=arr)
    val = Calc.get_num(sheet, f"{rng}{arr_size}")
    assert val == 3.14

    Lo.close(closeable=doc, deliver_ownership=False)


def test_set_array(loader) -> None:
    def arr_cb(row: int, col: int, prev_val) -> float:
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
    arr = TableHelper.to_2d_tuple(TableHelper.to_tuple(2))
    sheet = Calc.get_active_sheet(doc=doc)
    Calc.set_array(sheet=sheet, name="B2", values=arr)
    # GUI.set_visible(is_visible=True, odoc=doc)
    # Lo.delay(1500)
    val = Calc.get_num(sheet, "B2")
    assert val == 2.0

    arr_size = 12
    arr = TableHelper.to_2d_tuple(TableHelper.make_2d_array(arr_size, arr_size, arr_cb))
    rng = TableHelper.make_column_name(arr_size)
    # test
    # set_array(values: Sequence[Sequence[object]], sheet: XSpreadsheet, col_start: int, row_start: int, col_end:int, row_end: int)
    # keyword args
    col_start = 0
    row_start = 0
    col_end = arr_size - 1
    row_end = arr_size - 1
    Calc.set_array(values=arr, sheet=sheet, col_start=col_start, row_start=row_start, col_end=col_end, row_end=row_end)
    val = Calc.get_num(sheet, f"{rng}{arr_size}")
    assert val == float(arr_size * arr_size)

    # positiona args
    Calc.set_array(arr, sheet, col_start, row_start, col_end, row_end)
    val = Calc.get_num(sheet, f"{rng}{arr_size}")
    assert val == float(arr_size * arr_size)

    # def set_array(values: Sequence[Sequence[object]], cell_range: XCellRange)
    # keyword args
    cell_range = Calc.get_cell_range(
        sheet=sheet, start_col=col_start, start_row=row_start, end_col=col_end, end_row=row_end
    )
    Calc.set_array(values=arr, cell_range=cell_range)
    val = Calc.get_num(sheet, f"{rng}{arr_size}")
    assert val == float(arr_size * arr_size)
    # positional args
    cell_range = Calc.get_cell_range(sheet, col_start, row_start, col_end, row_end)
    Calc.set_array(arr, cell_range)
    val = Calc.get_num(sheet, f"{rng}{arr_size}")
    assert val == float(arr_size * arr_size)

    # set_array(values: Sequence[Sequence[object]], sheet: XSpreadsheet, addr: CellAddress)
    # keyword args
    addr = Calc.get_cell_address(sheet=sheet, cell_name="A1")
    Calc.set_array(values=arr, doc=doc, addr=addr)
    val = Calc.get_num(sheet, f"{rng}{arr_size}")
    assert val == float(arr_size * arr_size)

    # positional args
    Calc.set_array(arr, doc, addr)
    val = Calc.get_num(sheet, f"{rng}{arr_size}")
    assert val == float(arr_size * arr_size)

    with pytest.raises(TypeError):
        # error on unused key
        Calc.set_array(values=arr, cellRange=cell_range)
    with pytest.raises(TypeError):
        # error on incorrect number of args
        Calc.set_array(arr, doc, addr, 3)

    Lo.close(closeable=doc, deliver_ownership=False)


def test_get_array(loader) -> None:
    def arr_cb(row: int, col: int, prev_val) -> float:
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
    sheet = Calc.get_sheet(doc=doc, index=0)

    arr_size = 8
    arr = TableHelper.to_2d_tuple(TableHelper.make_2d_array(arr_size, arr_size, arr_cb))
    rng = TableHelper.make_column_name(arr_size)
    Calc.set_array(sheet=sheet, name=f"A1", values=arr)
    val = Calc.get_num(sheet, f"{rng}{arr_size}")
    assert val == float(arr_size * arr_size)

    # test overloads
    # get_array(sheet: XSpreadsheet, range_name: str)
    # by keyword
    rng_name = Calc.get_range_str(start_col=0, start_row=0, end_col=arr_size - 1, end_row=arr_size - 1)
    result_arr = Calc.get_array(sheet=sheet, range_name=rng_name)
    for row, row_data in enumerate(arr):
        for col, _ in enumerate(row_data):
            assert result_arr[row][col] == arr[row][col]
    # by position
    result_arr = Calc.get_array(sheet, rng_name)
    for row, row_data in enumerate(arr):
        for col, _ in enumerate(row_data):
            assert result_arr[row][col] == arr[row][col]

    # get_array(cell_range:XCellRange)
    # by Keyword
    cell_range = Calc.get_cell_range(sheet=sheet, range_name=rng_name)
    result_arr = Calc.get_array(cell_range=cell_range)
    for row, row_data in enumerate(arr):
        for col, _ in enumerate(row_data):
            assert result_arr[row][col] == arr[row][col]
    # by position
    result_arr = Calc.get_array(cell_range)
    for row, row_data in enumerate(arr):
        for col, _ in enumerate(row_data):
            assert result_arr[row][col] == arr[row][col]

    with pytest.raises(TypeError):
        # error on unused key
        Calc.get_array(sheet=sheet, rangeName=rng_name)
    with pytest.raises(TypeError):
        # error on incorrect number of args
        Calc.get_array(sheet, rng_name, 1)

    Lo.close(closeable=doc, deliver_ownership=False)


def test_print_array(capsys: pytest.CaptureFixture) -> None:
    from ooodev.office.calc import Calc

    Calc.print_array([])
    captured = capsys.readouterr()
    assert captured.out == "No data in array to print\n"
    arr1 = ((1, 2, 3), (4, 5, 6), (7, 8, 9))
    Calc.print_array(arr1)
    captured = capsys.readouterr()
    # assert captured.out == 'Row x Column size: 3 x 3\n1  2  3\n4  5  6\n7  8  9\n\n'
    assert (
        captured.out
        == """Row x Column size: 3 x 3
1  2  3
4  5  6
7  8  9

"""
    )


def test_get_float_array(loader) -> None:
    def arr_cb(row: int, col: int, prev_val) -> str:
        if row == 0 and col == 0:
            return "1"
        return str(int(prev_val) + 1)

    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    # from ooodev.utils.gui import GUI
    from ooodev.utils.gen_util import TableHelper

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    sheet = Calc.get_sheet(doc=doc, index=0)
    arr_size = 8
    arr = TableHelper.to_2d_tuple(TableHelper.make_2d_array(arr_size, arr_size, arr_cb))
    rng_name = Calc.get_range_str(start_col=0, start_row=0, end_col=arr_size - 1, end_row=arr_size - 1)
    Calc.set_array(values=arr, sheet=sheet, name=rng_name)
    arr_float = Calc.get_float_array(sheet=sheet, range_name=rng_name)
    for row, row_data in enumerate(arr):
        for col, _ in enumerate(row_data):
            assert arr_float[row][col] == float(arr[row][col])


# endregion set/get values in 2D array

# region    set/get rows and columns
def test_get_set_col(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    # from ooodev.utils.gui import GUI

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    sheet = Calc.get_sheet(doc=doc, index=0)

    vals_len = 12

    vals = [float(i) for i in range(vals_len)]
    # set_col(sheet: XSpreadsheet, values: Sequence[Any], cell_name: str)
    # keyword arguments
    Calc.set_col(sheet=sheet, values=vals, cell_name="A1")
    range_name = f"{Calc.get_cell_str(col=0, row=0)}:{Calc.get_cell_str(col=0, row= vals_len -1)}"
    new_vals = Calc.get_col(sheet=sheet, range_name=range_name)
    for i, val in enumerate(vals):
        assert new_vals[i] == val
    # Positional arguments
    Calc.set_col(sheet, vals, "A1")
    new_vals = Calc.get_col(sheet=sheet, range_name=range_name)
    for i, val in enumerate(vals):
        assert new_vals[i] == val
    # Mixed arguments
    Calc.set_col(sheet, vals, cell_name="A1")
    new_vals = Calc.get_col(sheet=sheet, range_name=range_name)
    for i, val in enumerate(vals):
        assert new_vals[i] == val

    # set_col(sheet: XSpreadsheet, values: Sequence[Any], col_start: int, row_start: int)
    # keyword arguments
    Calc.set_col(sheet=sheet, values=vals, col_start=0, row_start=0)
    new_vals = Calc.get_col(sheet=sheet, range_name=range_name)
    for i, val in enumerate(vals):
        assert new_vals[i] == val
    # Positional arguments
    Calc.set_col(sheet, vals, 0, 0)
    new_vals = Calc.get_col(sheet=sheet, range_name=range_name)
    for i, val in enumerate(vals):
        assert new_vals[i] == val

    with pytest.raises(TypeError):
        # error on unused key
        Calc.set_col(sheet, vals, cellName="A1")
    with pytest.raises(TypeError):
        # error on incorrect number of args
        Calc.set_col(sheet, vals)

    Lo.close(closeable=doc, deliver_ownership=False)


def test_get_set_row(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    # from ooodev.utils.gui import GUI

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    sheet = Calc.get_sheet(doc=doc, index=0)

    vals_len = 12

    vals = [float(i) for i in range(vals_len)]
    # set_col(sheet: XSpreadsheet, values: Sequence[Any], cell_name: str)
    # keyword arguments
    col = 2
    row = 2
    c_name = Calc.column_number_str(col)
    # overloads
    # set_row(sheet: XSpreadsheet, values: Sequence[Any], cell_name: str)
    # keyword
    Calc.set_row(sheet=sheet, values=vals, cell_name=f"{c_name}{col +1}")
    range_name = f"{Calc.get_cell_str(col=col, row=row)}:{Calc.get_cell_str(col=(vals_len -1) + col, row=row)}"
    new_vals = Calc.get_row(sheet=sheet, range_name=range_name)
    for i, val in enumerate(vals):
        assert new_vals[i] == val
    # positional
    Calc.set_row(sheet, vals, f"{c_name}{col +1}")
    new_vals = Calc.get_row(sheet=sheet, range_name=range_name)
    for i, val in enumerate(vals):
        assert new_vals[i] == val

    # set_row(sheet: XSpreadsheet, values: Sequence[Any], col_start: int, row_start: int)
    # keyword
    Calc.set_row(sheet=sheet, values=vals, col_start=col, row_start=row)
    new_vals = Calc.get_row(sheet=sheet, range_name=range_name)
    for i, val in enumerate(vals):
        assert new_vals[i] == val
    # positional
    Calc.set_row(sheet, vals, col, row)
    new_vals = Calc.get_row(sheet=sheet, range_name=range_name)
    for i, val in enumerate(vals):
        assert new_vals[i] == val

    with pytest.raises(TypeError):
        # error on unused key
        Calc.set_row(sheet=sheet, values=vals, cellName="test")
    with pytest.raises(TypeError):
        # error on incorrect number of args
        Calc.set_row(sheet, vals)

    # GUI.set_visible(is_visible=True, odoc=doc)
    # Lo.delay(3000)
    Lo.close(closeable=doc, deliver_ownership=False)


# endregion set/get rows and columns

# region    special cell types
def test_set_date(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    # from ooodev.utils.gui import GUI

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    sheet = Calc.get_sheet(doc=doc, index=0)

    Calc.set_date(sheet=sheet, cell_name="A1", day=22, month=11, year=2022)

    cell_str = Calc.get_string(sheet=sheet, cell_name="A1")
    assert cell_str == "44887.0"
    # GUI.set_visible(is_visible=True, odoc=doc)
    # Lo.delay(6000)
    Lo.close(closeable=doc, deliver_ownership=False)


def test_annotation(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    # from ooodev.utils.gui import GUI

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    sheet = Calc.get_sheet(doc=doc, index=0)
    cn = "B2"
    msg = "Hello World"
    ann = Calc.add_annotation(sheet=sheet, cell_name=cn, msg=msg)
    assert ann is not None
    ann = Calc.get_annotation(sheet=sheet, cell_name=cn)
    assert ann is not None
    ann_str = Calc.get_annotation_str(sheet=sheet, cell_name=cn)
    assert ann_str == msg
    # test getting of annotation with no annotation set
    ann = Calc.get_annotation(sheet=sheet, cell_name="G9")
    assert ann is not None
    #  # test getting of annotation string with no annotation set
    ann_str = Calc.get_annotation_str(sheet=sheet, cell_name="G9")
    assert ann_str == ""
    # GUI.set_visible(is_visible=True, odoc=doc)
    # Lo.delay(6000)
    Lo.close(closeable=doc, deliver_ownership=False)


# endregion special cell types

# region    get XCell and XCellRange methods
def test_get_cell(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    # from ooodev.utils.gui import GUI

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    sheet = Calc.get_sheet(doc=doc, index=0)
    test_val = "test"
    col = 2  # C
    row = 4
    addr = Calc.get_cell_address(sheet=sheet, col=col, row=row)
    cell_name = Calc.get_cell_str(addr=addr)

    cell_range = Calc.get_cell_range(sheet=sheet, start_col=col, start_row=row, end_col=col, end_row=row)
    Calc.set_val(value=test_val, sheet=sheet, cell_name=cell_name)
    # get_cell(sheet: XSpreadsheet, addr: CellAddress)
    cell = Calc.get_cell(sheet=sheet, addr=addr)
    assert cell is not None
    cell = Calc.get_cell(sheet, addr)
    assert cell is not None
    val = Calc.get_string(cell)
    assert val == test_val

    # get_cell(sheet: XSpreadsheet, cell_name: str)
    cell = Calc.get_cell(sheet=sheet, cell_name=cell_name)
    assert cell is not None
    cell = Calc.get_cell(sheet, cell_name)
    assert cell is not None
    val = Calc.get_string(cell)
    assert val == test_val

    # get_cell(sheet: XSpreadsheet, column: int, row: int
    cell = Calc.get_cell(sheet=sheet, col=col, row=row)
    assert cell is not None
    val = Calc.get_string(cell)
    assert val == test_val
    cell = Calc.get_cell(sheet, col, row)
    assert cell is not None

    #  get_cell(cell_range: XCellRange)
    cell = Calc.get_cell(cell_range=cell_range)
    val = Calc.get_string(cell)
    assert cell is not None
    val = Calc.get_string(cell)
    assert val == test_val
    assert val == test_val
    cell = Calc.get_cell(cell_range)
    assert cell is not None

    #  get_cell(cell_range: XCellRange, column: int, row: int)
    # cell range is relative position.
    # if a range is C4:E9 then Cell range at col=0 ,row=0 is C4
    cell = Calc.get_cell(cell_range=cell_range, col=0, row=0)
    assert cell is not None
    val = Calc.get_string(cell)
    assert val == test_val
    cell = Calc.get_cell(cell_range, 0, 0)
    assert cell is not None

    with pytest.raises(TypeError):
        # error on unused key
        Calc.get_cell(sheet=sheet, cellName="A2")
    with pytest.raises(TypeError):
        # error on incorrect number of args
        Calc.get_cell()

    Lo.close(closeable=doc, deliver_ownership=False)


def test_is_single_cell_range(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    sheet = Calc.get_sheet(doc=doc, index=0)
    cr_addr = Calc.get_address(sheet=sheet, range_name="A1:A1")
    assert Calc.is_single_cell_range(cr_addr)
    cr_addr = Calc.get_address(sheet=sheet, range_name="A1:B1")
    assert Calc.is_single_cell_range(cr_addr) == False
    Lo.close(closeable=doc, deliver_ownership=False)


def test_get_cell_range(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    sheet = Calc.get_sheet(doc=doc, index=0)
    single_col_start = 1
    single_row_start = 1
    single_col_end = 1
    single_row_end = 1

    multi_row_start = 2
    multi_col_start = 2
    multi_col_end = 5
    multi_row_end = 21

    rng_single = Calc.get_range_str(
        start_col=single_col_start, start_row=single_row_start, end_col=single_col_end, end_row=single_row_end
    )

    rng_multi = Calc.get_range_str(
        start_col=multi_col_start, start_row=multi_row_start, end_col=multi_col_end, end_row=multi_row_end
    )
    cr_addr_single = Calc.get_address(sheet=sheet, range_name=rng_single)
    cr_addr_muilti = Calc.get_address(sheet=sheet, range_name=rng_multi)

    # get_cell_range(sheet: XSpreadsheet, cr_addr: CellRangeAddress)
    rng = Calc.get_cell_range(sheet=sheet, cr_addr=cr_addr_single)
    assert rng is not None
    addr = Calc.get_address(cell_range=rng)
    assert Calc.is_single_cell_range(addr)
    rng = Calc.get_cell_range(sheet, cr_addr_single)
    assert rng is not None
    addr = Calc.get_address(cell_range=rng)
    assert Calc.is_single_cell_range(addr)
    rng = Calc.get_cell_range(sheet, cr_addr_muilti)
    assert rng is not None
    addr = Calc.get_address(cell_range=rng)
    assert Calc.is_single_cell_range(addr) == False

    # get_cell_range(sheet: XSpreadsheet, range_name: str)
    rng = Calc.get_cell_range(sheet=sheet, range_name=rng_single)
    assert rng is not None
    addr = Calc.get_address(cell_range=rng)
    assert Calc.is_single_cell_range(addr)
    rng = Calc.get_cell_range(sheet, rng_single)
    assert rng is not None
    addr = Calc.get_address(cell_range=rng)
    assert Calc.is_single_cell_range(addr)
    rng = Calc.get_cell_range(sheet, rng_multi)
    assert rng is not None
    addr = Calc.get_address(cell_range=rng)
    assert Calc.is_single_cell_range(addr) == False

    #  get_cell_range(sheet: XSpreadsheet, col_start: int, row_start: int, col_end: int, row_end: int)
    rng = Calc.get_cell_range(
        sheet=sheet,
        start_col=single_col_start,
        start_row=single_row_start,
        end_col=single_col_end,
        end_row=single_row_end,
    )
    assert rng is not None
    addr = Calc.get_address(cell_range=rng)
    assert Calc.is_single_cell_range(addr)
    rng = Calc.get_cell_range(sheet, single_col_start, single_row_start, single_col_end, single_row_end)
    assert rng is not None
    addr = Calc.get_address(cell_range=rng)
    assert Calc.is_single_cell_range(addr)
    rng = Calc.get_cell_range(sheet, multi_col_start, multi_row_start, multi_col_end, multi_row_end)
    assert rng is not None
    addr = Calc.get_address(cell_range=rng)
    assert Calc.is_single_cell_range(addr) == False

    with pytest.raises(TypeError):
        # error on unused key
        Calc.get_cell_range(sheet=sheet, rangeName="A2:B5")
    with pytest.raises(TypeError):
        # error on incorrect number of args
        Calc.get_cell_range(sheet)

    Lo.close(closeable=doc, deliver_ownership=False)


def test_find_used_range(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    sheet = Calc.get_sheet(doc=doc, index=0)

    # find_used_range(sheet: XSpreadsheet)
    rng = Calc.find_used_range(sheet=sheet)
    assert rng is not None
    rng_str = Calc.get_range_str(cell_range=rng)
    assert rng_str == "A1:A1"

    rng = Calc.find_used_range(sheet=sheet, cell_name="C2")
    assert rng is not None
    rng_str = Calc.get_range_str(cell_range=rng)
    assert rng_str == "A1:A1"

    Lo.close(closeable=doc, deliver_ownership=False)


def test_get_col_range(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    sheet = Calc.get_sheet(doc=doc, index=0)
    test_val = "test"
    index = 3
    Calc.set_val(value=test_val, sheet=sheet, col=3, row=0)
    rng = Calc.get_col_range(sheet=sheet, idx=index)
    addr = Calc.get_address(cell_range=rng)
    assert Calc.is_single_cell_range(addr) == False
    cell = Calc.get_cell(cell_range=rng)
    val = Calc.get_string(cell)
    assert val == test_val
    Lo.close(closeable=doc, deliver_ownership=False)


def test_get_row_range(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    sheet = Calc.get_sheet(doc=doc, index=0)
    test_val = "test"
    index = 3
    Calc.set_val(value=test_val, sheet=sheet, col=0, row=index)
    rng = Calc.get_row_range(sheet=sheet, idx=index)
    addr = Calc.get_address(cell_range=rng)
    assert Calc.is_single_cell_range(addr) == False
    cell = Calc.get_cell(cell_range=rng)
    val = Calc.get_string(cell)
    assert val == test_val
    Lo.close(closeable=doc, deliver_ownership=False)


# endregion get XCell and XCellRange methods

# region    convert cell/cellrange names to positions


def test_get_cell_range_positions() -> None:
    from ooodev.office.calc import Calc

    points = Calc.get_cell_range_positions(range_name="A1:C5")
    assert len(points) == 2
    assert points[0].X == 0  # A
    assert points[0].Y == 0  # 1
    assert points[1].X == 2  # C
    assert points[1].Y == 4  # 5


def test_get_cell_position() -> None:
    from ooodev.office.calc import Calc

    # get_cell_position(cell_name: str)
    p = Calc.get_cell_position(cell_name="C5")
    assert p.X == 2
    assert p.Y == 4
    p = Calc.get_cell_position("C5")
    assert p.X == 2
    assert p.Y == 4


def test_get_cell_pos(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    sheet = Calc.get_sheet(doc=doc, index=0)

    p1 = Calc.get_cell_pos(sheet=sheet, cell_name="C5")
    p2 = Calc.get_cell_pos(sheet=sheet, cell_name="A1")
    assert p1.X != p2.X
    assert p1.Y != p2.X
    Lo.close(closeable=doc, deliver_ownership=False)


def test_column_number_str() -> None:
    from ooodev.office.calc import Calc

    s = Calc.column_number_str(0)
    assert s == "A"

    s = Calc.column_number_str(99)
    assert s == "CV"


def test_column_string_to_number() -> None:
    from ooodev.office.calc import Calc

    i = Calc.column_string_to_number("A")
    assert i == 0

    i = Calc.column_string_to_number("A4")
    assert i == 0

    i = Calc.column_string_to_number("CV")
    assert i == 99

    i = Calc.column_string_to_number("CV22")
    assert i == 99


def test_row_string_to_number() -> None:
    from ooodev.office.calc import Calc

    i = Calc.row_string_to_number("4")
    assert i == 3

    i = Calc.row_string_to_number("Bc4")
    assert i == 3

    i = Calc.row_string_to_number("DR8")
    assert i == 7

    # negative or invalid is returned as 0
    i = Calc.row_string_to_number("-4")
    assert i == 0
    i = Calc.row_string_to_number("BA-4")
    assert i == 0


# endregion convert cell/cellrange names to positions

# region    get cell and cell range addresses
def test_get_cell_address(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    sheet = Calc.get_sheet(doc=doc, index=0)

    cell_name = "C2"
    cell = Calc.get_cell(sheet=sheet, cell_name=cell_name)

    #  get_cell_address(cell: XCell)
    addr = Calc.get_cell_address(cell=cell)
    assert addr.Column == 2
    assert addr.Row == 1
    addr = Calc.get_cell_address(cell)
    assert addr.Column == 2
    assert addr.Row == 1

    # get_cell_address(sheet: XSpreadsheet, cell_name: str)
    addr = Calc.get_cell_address(sheet=sheet, cell_name=cell_name)
    assert addr.Column == 2
    assert addr.Row == 1
    addr = Calc.get_cell_address(sheet, cell_name)
    assert addr.Column == 2
    assert addr.Row == 1

    # get_cell_address(sheet: XSpreadsheet, col: int, row: int)
    addr = Calc.get_cell_address(sheet=sheet, col=2, row=1)
    assert addr.Column == 2
    assert addr.Row == 1
    addr = Calc.get_cell_address(sheet, 2, 1)
    assert addr.Column == 2
    assert addr.Row == 1

    # get_cell_address(sheet: XSpreadsheet, addr: CellAddress)
    addr2 = Calc.get_cell_address(sheet=sheet, addr=addr)
    assert addr2.Column == 2
    assert addr2.Row == 1
    addr2 = Calc.get_cell_address(sheet, addr)
    assert addr2.Column == 2
    assert addr2.Row == 1

    with pytest.raises(TypeError):
        # error on unused key
        Calc.get_cell_address(sheet=sheet, cellName="B4")
    with pytest.raises(TypeError):
        # error on incorrect number of args
        Calc.get_cell_address(sheet, 2, 1, 4)

    Lo.close(closeable=doc, deliver_ownership=False)


def test_get_address(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    sheet = Calc.get_sheet(doc=doc, index=0)

    start_col = 2
    start_row = 2
    end_col = 5
    end_row = 21

    rng = Calc.get_cell_range(sheet=sheet, start_col=start_col, start_row=start_row, end_col=end_col, end_row=end_row)
    range_name = Calc.get_range_str(cell_range=rng)
    # get_address(cell_range: XCellRange)
    cr_addr = Calc.get_address(cell_range=rng)
    assert cr_addr.StartColumn == start_col
    assert cr_addr.StartRow == start_row
    assert cr_addr.EndColumn == end_col
    assert cr_addr.EndRow == end_row
    cr_addr = Calc.get_address(rng)
    assert cr_addr.StartColumn == start_col
    assert cr_addr.StartRow == start_row
    assert cr_addr.EndColumn == end_col
    assert cr_addr.EndRow == end_row

    # get_address(sheet: XSpreadsheet, range_name: str)
    cr_addr = Calc.get_address(sheet=sheet, range_name=range_name)
    assert cr_addr.StartColumn == start_col
    assert cr_addr.StartRow == start_row
    assert cr_addr.EndColumn == end_col
    assert cr_addr.EndRow == end_row
    cr_addr = Calc.get_address(sheet, range_name)
    assert cr_addr.StartColumn == start_col
    assert cr_addr.StartRow == start_row
    assert cr_addr.EndColumn == end_col
    assert cr_addr.EndRow == end_row

    # get_address(sheet: XSpreadsheet, start_col: int, start_row: int, end_col: int, end_row: int)
    cr_addr = Calc.get_address(sheet=sheet, start_col=start_col, start_row=start_row, end_col=end_col, end_row=end_row)
    assert cr_addr.StartColumn == start_col
    assert cr_addr.StartRow == start_row
    assert cr_addr.EndColumn == end_col
    assert cr_addr.EndRow == end_row
    cr_addr = Calc.get_address(sheet, start_col, start_row, end_col, end_row)
    assert cr_addr.StartColumn == start_col
    assert cr_addr.StartRow == start_row
    assert cr_addr.EndColumn == end_col
    assert cr_addr.EndRow == end_row

    with pytest.raises(TypeError):
        # error on unused key
        Calc.get_address(cellRange=rng)
    with pytest.raises(TypeError):
        # error on incorrect number of args
        Calc.get_address(sheet, start_col, start_row)

    Lo.close(closeable=doc, deliver_ownership=False)


def test_print_cell_address(capsys: pytest.CaptureFixture, loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    sheet = Calc.get_sheet(doc=doc, index=0)
    print_result = "Cell: Sheet1.D3\n"
    col = 3
    row = 2
    addr = Calc.get_cell_address(sheet=sheet, col=col, row=row)
    cell = Calc.get_cell(sheet=sheet, addr=addr)
    # print_cell_address(cell: XCell)
    captured = capsys.readouterr()  # clear buffer
    Calc.print_cell_address(cell=cell)
    captured = capsys.readouterr()
    assert captured.out == print_result
    Calc.print_cell_address(cell)
    captured = capsys.readouterr()
    assert captured.out == print_result

    # print_cell_address(addr: CellAddress)
    Calc.print_cell_address(addr=addr)
    captured = capsys.readouterr()
    assert captured.out == print_result
    Calc.print_cell_address(addr)
    captured = capsys.readouterr()
    assert captured.out == print_result

    with pytest.raises(TypeError):
        # error on unused key
        Calc.print_cell_address(cel=cell)
    with pytest.raises(TypeError):
        # error on incorrect number of args
        Calc.print_cell_address(1, 2)
    Lo.close(closeable=doc, deliver_ownership=False)


def test_print_address(capsys: pytest.CaptureFixture, loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    sheet = Calc.get_sheet(doc=doc, index=0)
    print_result = "Range: Sheet1.C3:F22\n"
    start_col = 2
    start_row = 2
    end_col = 5
    end_row = 21

    rng = Calc.get_cell_range(sheet=sheet, start_col=start_col, start_row=start_row, end_col=end_col, end_row=end_row)
    cr_addr = Calc.get_address(cell_range=rng)

    # print_address(cell_range: XCellRange)
    captured = capsys.readouterr()  # clear buffer
    Calc.print_address(cell_range=rng)
    captured = capsys.readouterr()
    assert captured.out == print_result
    Calc.print_address(rng)
    captured = capsys.readouterr()
    assert captured.out == print_result

    # print_address(cr_addr: CellRangeAddress)
    Calc.print_address(cr_addr=cr_addr)
    captured = capsys.readouterr()
    assert captured.out == print_result
    Calc.print_address(cr_addr)
    captured = capsys.readouterr()
    assert captured.out == print_result

    with pytest.raises(TypeError):
        # error on unused key
        Calc.print_address(cellRange=rng)
    with pytest.raises(TypeError):
        # error on incorrect number of args
        Calc.print_address(1, 2)

    Lo.close(closeable=doc, deliver_ownership=False)


def test_print_addresses(capsys: pytest.CaptureFixture, loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    sheet = Calc.get_sheet(doc=doc, index=0)
    print_result = """No of cellrange addresses: 2
Range: Sheet1.C3:F22
Range: Sheet1.B26:J45

"""
    first_start_col = 2
    first_start_row = 2
    first_end_col = 5
    first_end_row = 21
    first_rng = Calc.get_cell_range(
        sheet=sheet, start_col=first_start_col, start_row=first_start_row, end_col=first_end_col, end_row=first_end_row
    )
    first_cr_addr = Calc.get_address(cell_range=first_rng)
    second_start_col = 1
    second_start_row = 25
    second_end_col = 9
    second_end_row = 44
    second_rng = Calc.get_cell_range(
        sheet=sheet,
        start_col=second_start_col,
        start_row=second_start_row,
        end_col=second_end_col,
        end_row=second_end_row,
    )
    second_cr_addr = Calc.get_address(cell_range=second_rng)
    captured = capsys.readouterr()  # clear buffer
    Calc.print_addresses(first_cr_addr, second_cr_addr)
    captured = capsys.readouterr()
    assert captured.out == print_result
    Lo.close(closeable=doc, deliver_ownership=False)


def test_get_cell_series(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc
    from com.sun.star.sheet import XCellSeries

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    sheet = Calc.get_sheet(doc=doc, index=0)
    series = Calc.get_cell_series(sheet=sheet, range_name="A2:B6")
    assert Lo.qi(XCellSeries, series) is not None
    Lo.close(closeable=doc, deliver_ownership=False)


def test_is_equal_addresses(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    sheet = Calc.get_sheet(doc=doc, index=0)

    first_start_col = 2
    first_start_row = 2
    first_end_col = 5
    first_end_row = 21
    first_rng = Calc.get_cell_range(
        sheet=sheet, start_col=first_start_col, start_row=first_start_row, end_col=first_end_col, end_row=first_end_row
    )
    first_cr_addr = Calc.get_address(cell_range=first_rng)
    second_start_col = 1
    second_start_row = 25
    second_end_col = 9
    second_end_row = 44
    second_rng = Calc.get_cell_range(
        sheet=sheet,
        start_col=second_start_col,
        start_row=second_start_row,
        end_col=second_end_col,
        end_row=second_end_row,
    )
    first_addr = Calc.get_cell_address(sheet=sheet, col=first_start_col, row=first_start_row)
    second_addr = Calc.get_cell_address(sheet=sheet, col=second_start_col, row=second_start_row)

    # is_equal_addresses(addr1: CellAddress, addr2: CellAddress)
    assert Calc.is_equal_addresses(addr1=first_addr, addr2=second_addr) == False
    assert Calc.is_equal_addresses(first_addr, second_addr) == False
    assert Calc.is_equal_addresses(addr1=first_addr, addr2=first_addr)
    assert Calc.is_equal_addresses(first_addr, first_addr)

    # is_equal_addresses(addr1: CellRangeAddress, addr2: CellRangeAddress)
    assert Calc.is_equal_addresses(addr1=first_cr_addr, addr2=second_rng) == False
    assert Calc.is_equal_addresses(first_cr_addr, second_rng) == False
    assert Calc.is_equal_addresses(addr1=first_cr_addr, addr2=first_cr_addr)
    assert Calc.is_equal_addresses(first_cr_addr, first_cr_addr)

    # test missmatched
    assert Calc.is_equal_addresses(addr1=first_cr_addr, addr2=second_addr) == False
    assert Calc.is_equal_addresses(first_cr_addr, second_addr) == False

    # test bad input
    assert Calc.is_equal_addresses(first_cr_addr, object()) == False

    Lo.close(closeable=doc, deliver_ownership=False)


# endregion get cell and cell range addresses

# region    convert cell range address to string
def test_get_range_str(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    sheet = Calc.get_sheet(doc=doc, index=0)
    sheet_name = Calc.get_sheet_name(sheet=sheet)

    start_col = 2
    start_row = 2
    end_col = 5
    end_row = 21
    rng_str = f"{Calc.get_cell_str(col=start_col, row=start_row)}:{Calc.get_cell_str(col=end_col, row=end_row)}"
    sheet_rng_str = f"{sheet_name}.{rng_str}"
    rng = Calc.get_cell_range(sheet=sheet, start_col=start_col, start_row=start_row, end_col=end_col, end_row=end_row)
    cr_addr = Calc.get_address(cell_range=rng)

    # get_range_str(cell_range: XCellRange)
    result = Calc.get_range_str(cell_range=rng)
    assert result == rng_str
    result = Calc.get_range_str(rng)
    assert result == rng_str

    # get_range_str(cr_addr: CellRangeAddress)
    result = Calc.get_range_str(cr_addr=cr_addr)
    assert result == rng_str
    result = Calc.get_range_str(cr_addr)
    assert result == rng_str

    # get_range_str(cell_range: XCellRange, sheet: XSpreadsheet)
    result = Calc.get_range_str(cell_range=rng, sheet=sheet)
    assert result == sheet_rng_str
    result = Calc.get_range_str(rng, sheet)
    assert result == sheet_rng_str

    # get_range_str(cr_addr: CellRangeAddress, sheet: XSpreadsheet)
    result = Calc.get_range_str(cr_addr=cr_addr, sheet=sheet)
    assert result == sheet_rng_str
    result = Calc.get_range_str(cr_addr, sheet)
    assert result == sheet_rng_str

    # get_range_str(start_col: int, start_row: int, end_col: int, end_row: int)
    result = Calc.get_range_str(start_col=start_col, start_row=start_row, end_col=end_col, end_row=end_row)
    assert result == rng_str
    result = Calc.get_range_str(start_col, start_row, end_col, end_row)
    assert result == rng_str

    # get_range_str(start_col: int, start_row: int, end_col: int, end_row: int, sheet: XSpreadsheet)
    result = Calc.get_range_str(
        start_col=start_col, start_row=start_row, end_col=end_col, end_row=end_row, sheet=sheet
    )
    assert result == sheet_rng_str
    result = Calc.get_range_str(start_col, start_row, end_col, end_row, sheet)
    assert result == sheet_rng_str

    with pytest.raises(TypeError):
        # error on unused key
        Calc.get_range_str(cellRange="A2:B5")
    with pytest.raises(TypeError):
        # error on incorrect number of args
        Calc.get_range_str(cr_addr, sheet, 2)

    Lo.close(closeable=doc, deliver_ownership=False)


def test_get_cell_str(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    sheet = Calc.get_sheet(doc=doc, index=0)
    col = 2
    row = 3
    name = f"{Calc.column_number_str(col=col)}{row+1}"
    addr = Calc.get_cell_address(sheet=sheet, col=col, row=row)
    cell = Calc.get_cell(sheet=sheet, addr=addr)

    # get_cell_str(addr: CellAddress)
    result = Calc.get_cell_str(addr=addr)
    assert result == name
    result = Calc.get_cell_str(addr)
    assert result == name

    # get_cell_str(cell: XCell)
    result = Calc.get_cell_str(cell=cell)
    assert result == name
    result = Calc.get_cell_str(cell)
    assert result == name

    # get_cell_str(col: int, row: int)
    result = Calc.get_cell_str(col=col, row=row)
    assert result == name
    result = Calc.get_cell_str(col, row)
    assert result == name

    with pytest.raises(TypeError):
        # error on unused key
        Calc.get_cell_str(adr=addr)
    with pytest.raises(TypeError):
        # error on incorrect number of args
        Calc.get_cell_str(addr, sheet)

    Lo.close(closeable=doc, deliver_ownership=False)


# endregion convert cell range address to string

# region    search
def test_find_all(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc
    from com.sun.star.util import XSearchable

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    sheet = Calc.get_sheet(doc=doc, index=0)
    test_val = "test"
    Calc.set_val(value=test_val, sheet=sheet, cell_name="A1")
    Calc.set_val(value=test_val, sheet=sheet, cell_name="C3")
    srch = Lo.qi(XSearchable, sheet)
    sd = srch.createSearchDescriptor()
    sd.setSearchString(test_val)
    results = Calc.find_all(srch=srch, sd=sd)
    assert len(results) == 2
    assert Calc.get_range_str(cell_range=results[0]) == "A1:A1"
    assert Calc.get_range_str(cell_range=results[1]) == "C3:C3"

    sd.setSearchString("hello")
    results = Calc.find_all(srch=srch, sd=sd)
    assert results is None
    Lo.close(closeable=doc, deliver_ownership=False)


# endregion search

# region    cell decoration
def test_create_cell_style(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    style = Calc.create_cell_style(doc=doc, style_name="Fancy")
    assert style is not None
    assert style.getName() == "Fancy"

    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    sheet = Calc.get_sheet(doc=doc, index=0)
    start_col = 2
    start_row = 2
    end_col = 5
    end_row = 21
    rng_str = Calc.get_range_str(start_col=start_col, start_row=start_row, end_col=end_col, end_row=end_row)
    cell_range = Calc.get_cell_range(
        sheet=sheet, start_col=start_col, start_row=start_row, end_col=end_col, end_row=end_row
    )
    style = "Accent 1"
    # change_style(sheet: XSpreadsheet, style_name: str, range_name: str)
    result = Calc.change_style(sheet=sheet, style_name=style, range_name=rng_str)
    assert result
    result = Calc.change_style(sheet, style, rng_str)
    assert result

    # change_style(sheet: XSpreadsheet, style_name: str, cell_range: XCellRange)
    result = Calc.change_style(sheet=sheet, style_name=style, cell_range=cell_range)
    assert result
    result = Calc.change_style(sheet, style, cell_range)
    assert result

    # change_style(sheet: XSpreadsheet, style_name: str, start_col: int, start_row: int, end_col: int, end_row: int)
    result = Calc.change_style(
        sheet=sheet, style_name=style, start_col=start_col, start_row=start_row, end_col=end_col, end_row=end_row
    )
    assert result
    result = Calc.change_style(sheet, style, start_col, start_row, end_col, end_row)
    assert result

    style = "Non existing style"
    # change_style(sheet: XSpreadsheet, style_name: str, range_name: str)
    result = Calc.change_style(sheet=sheet, style_name=style, range_name=rng_str)
    assert result == False

    with pytest.raises(TypeError):
        # error on unused key
        Calc.change_style(sheet=sheet, style_name=style, cellRange=cell_range)
    with pytest.raises(TypeError):
        # error on incorrect number of args
        Calc.change_style(sheet, style, cell_range, rng_str)

    Lo.close(closeable=doc, deliver_ownership=False)


def test_add_remove_border(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc
    from ooodev.utils.props import Props
    from ooodev.utils import color
    from com.sun.star.table import TableBorder2
    from ooodev.utils.gui import GUI

    visible = False
    delay = 0 # 300
    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    sheet = Calc.get_sheet(doc=doc, index=0)
    try:
        GUI.set_visible(is_visible=visible, odoc=doc)
        Lo.delay(delay)
        rng_name = "B1:F8"
        # add_border(sheet: XSpreadsheet, range_name: str)
        rng = Calc.add_border(sheet=sheet, range_name=rng_name)
        assert rng is not None
        tbl_border = cast(TableBorder2, Props.get_property(prop_set=rng, name="TableBorder2"))
        assert tbl_border.LeftLine.Color == 0
        assert tbl_border.RightLine.Color == 0
        assert tbl_border.TopLine.Color == 0
        assert tbl_border.BottomLine.Color == 0
        Lo.delay(delay)

        rng = Calc.remove_border(sheet=sheet, range_name=rng_name)
        assert rng is not None
        tbl_border = cast(TableBorder2, Props.get_property(prop_set=rng, name="TableBorder2"))
        assert tbl_border.LeftLine.Color == 0
        assert tbl_border.RightLine.Color == 0
        assert tbl_border.TopLine.Color == 0
        assert tbl_border.BottomLine.Color == 0
        assert tbl_border.BottomLine.LineWidth == 0
        Lo.delay(delay)

        cell_rng = Calc.add_border(sheet, rng_name)
        assert cell_rng is not None
        tbl_border = cast(TableBorder2, Props.get_property(prop_set=cell_rng, name="TableBorder2"))
        assert tbl_border.LeftLine.Color == 0
        assert tbl_border.RightLine.Color == 0
        assert tbl_border.TopLine.Color == 0
        assert tbl_border.BottomLine.Color == 0
        Lo.delay(delay)

        rng = Calc.remove_border(sheet, rng_name)
        assert rng is not None
        tbl_border = cast(TableBorder2, Props.get_property(prop_set=rng, name="TableBorder2"))
        assert tbl_border.LeftLine.Color == 0
        assert tbl_border.RightLine.Color == 0
        assert tbl_border.TopLine.Color == 0
        assert tbl_border.BottomLine.Color == 0
        assert tbl_border.BottomLine.LineWidth == 0
        Lo.delay(delay)

        # add_border(sheet: XSpreadsheet, cell_range: XCellRange)
        Lo.delay(delay)
        cell_rng = Calc.add_border(sheet=sheet, cell_range=rng)
        assert cell_rng is not None
        tbl_border = cast(TableBorder2, Props.get_property(prop_set=cell_rng, name="TableBorder2"))
        assert tbl_border.LeftLine.Color == 0
        assert tbl_border.RightLine.Color == 0
        assert tbl_border.TopLine.Color == 0
        assert tbl_border.BottomLine.Color == 0
        Lo.delay(delay)

        rng = Calc.remove_border(sheet=sheet, cell_range=rng)
        assert rng is not None
        tbl_border = cast(TableBorder2, Props.get_property(prop_set=rng, name="TableBorder2"))
        assert tbl_border.LeftLine.Color == 0
        assert tbl_border.RightLine.Color == 0
        assert tbl_border.TopLine.Color == 0
        assert tbl_border.BottomLine.Color == 0
        assert tbl_border.BottomLine.LineWidth == 0
        Lo.delay(delay)

        cell_rng = Calc.add_border(sheet, rng)
        assert cell_rng is not None
        tbl_border = cast(TableBorder2, Props.get_property(prop_set=cell_rng, name="TableBorder2"))
        assert tbl_border.LeftLine.Color == 0
        assert tbl_border.RightLine.Color == 0
        assert tbl_border.TopLine.Color == 0
        assert tbl_border.BottomLine.Color == 0
        Lo.delay(delay)

        rng = Calc.remove_border(sheet, rng)
        assert rng is not None
        tbl_border = cast(TableBorder2, Props.get_property(prop_set=rng, name="TableBorder2"))
        assert tbl_border.LeftLine.Color == 0
        assert tbl_border.RightLine.Color == 0
        assert tbl_border.TopLine.Color == 0
        assert tbl_border.BottomLine.Color == 0
        assert tbl_border.BottomLine.LineWidth == 0
        Lo.delay(delay)

        # add_border(sheet: XSpreadsheet, cell_range: XCellRange, color: int)
        cell_rng = Calc.add_border(sheet=sheet, cell_range=rng, color=color.CommonColor.GREEN)
        assert cell_rng is not None
        tbl_border = cast(TableBorder2, Props.get_property(prop_set=cell_rng, name="TableBorder2"))
        assert tbl_border.LeftLine.Color == color.CommonColor.GREEN
        assert tbl_border.RightLine.Color == color.CommonColor.GREEN
        assert tbl_border.TopLine.Color == color.CommonColor.GREEN
        assert tbl_border.BottomLine.Color == color.CommonColor.GREEN
        Lo.delay(delay)

        rng = Calc.remove_border(sheet=sheet, cell_range=rng)
        assert rng is not None
        tbl_border = cast(TableBorder2, Props.get_property(prop_set=rng, name="TableBorder2"))
        assert tbl_border.LeftLine.Color == 0
        assert tbl_border.RightLine.Color == 0
        assert tbl_border.TopLine.Color == 0
        assert tbl_border.BottomLine.Color == 0
        assert tbl_border.BottomLine.LineWidth == 0
        Lo.delay(delay)

        cell_rng = Calc.add_border(sheet, rng, color.CommonColor.DARK_BLUE)
        assert cell_rng is not None
        tbl_border = cast(TableBorder2, Props.get_property(prop_set=cell_rng, name="TableBorder2"))
        assert tbl_border.LeftLine.Color == color.CommonColor.DARK_BLUE
        assert tbl_border.RightLine.Color == color.CommonColor.DARK_BLUE
        assert tbl_border.TopLine.Color == color.CommonColor.DARK_BLUE
        assert tbl_border.BottomLine.Color == color.CommonColor.DARK_BLUE
        Lo.delay(delay)

        rng = Calc.remove_border(sheet=sheet, cell_range=rng)
        assert rng is not None
        tbl_border = cast(TableBorder2, Props.get_property(prop_set=rng, name="TableBorder2"))
        assert tbl_border.LeftLine.Color == 0
        assert tbl_border.RightLine.Color == 0
        assert tbl_border.TopLine.Color == 0
        assert tbl_border.BottomLine.Color == 0
        assert tbl_border.BottomLine.LineWidth == 0
        Lo.delay(delay)

        # add_border(sheet: XSpreadsheet, range_name: str, color: int)
        cell_rng = Calc.add_border(sheet=sheet, range_name=rng_name, color=color.CommonColor.GREEN)
        assert cell_rng is not None
        tbl_border = cast(TableBorder2, Props.get_property(prop_set=cell_rng, name="TableBorder2"))
        assert tbl_border.LeftLine.Color == color.CommonColor.GREEN
        assert tbl_border.RightLine.Color == color.CommonColor.GREEN
        assert tbl_border.TopLine.Color == color.CommonColor.GREEN
        assert tbl_border.BottomLine.Color == color.CommonColor.GREEN
        Lo.delay(delay)

        rng = Calc.remove_border(
            sheet,
            rng,
            Calc.BorderEnum.BOTTOM_BORDER
            | Calc.BorderEnum.LEFT_BORDER
            | Calc.BorderEnum.RIGHT_BORDER
            | Calc.BorderEnum.TOP_BORDER,
        )
        assert rng is not None
        tbl_border = cast(TableBorder2, Props.get_property(prop_set=rng, name="TableBorder2"))
        assert tbl_border.LeftLine.Color == 0
        assert tbl_border.RightLine.Color == 0
        assert tbl_border.TopLine.Color == 0
        assert tbl_border.BottomLine.Color == 0
        assert tbl_border.BottomLine.LineWidth == 0
        Lo.delay(delay)

        cell_rng = Calc.add_border(sheet, rng, color.CommonColor.DARK_BLUE)
        assert cell_rng is not None
        tbl_border = cast(TableBorder2, Props.get_property(prop_set=cell_rng, name="TableBorder2"))
        assert tbl_border.LeftLine.Color == color.CommonColor.DARK_BLUE
        assert tbl_border.RightLine.Color == color.CommonColor.DARK_BLUE
        assert tbl_border.TopLine.Color == color.CommonColor.DARK_BLUE
        assert tbl_border.BottomLine.Color == color.CommonColor.DARK_BLUE
        Lo.delay(delay)

        Calc.remove_border(sheet=sheet, cell_range=rng)
        Lo.delay(delay)

        # add_border(sheet: XSpreadsheet, cell_range: XCellRange, color: int, border_vals: int)
        bval = Calc.BorderEnum.TOP_BORDER | Calc.BorderEnum.LEFT_BORDER
        cell_rng = Calc.add_border(sheet=sheet, cell_range=rng, color=color.CommonColor.BLACK)  # reset colors
        cell_rng = Calc.add_border(sheet=sheet, cell_range=rng, color=color.CommonColor.GREEN, border_vals=int(bval))
        tbl_border = cast(TableBorder2, Props.get_property(prop_set=cell_rng, name="TableBorder2"))
        assert tbl_border.TopLine.Color == color.CommonColor.GREEN
        assert tbl_border.LeftLine.Color == color.CommonColor.GREEN
        assert tbl_border.RightLine.Color == color.CommonColor.BLACK
        assert tbl_border.BottomLine.Color == color.CommonColor.BLACK
        Lo.delay(delay)

        Calc.remove_border(sheet=sheet, cell_range=rng)
        Lo.delay(delay)

        cell_rng = Calc.add_border(sheet=sheet, cell_range=rng, color=color.CommonColor.BLACK)  # reset colors
        cell_rng = Calc.add_border(sheet, rng, color.CommonColor.DARK_BLUE, int(bval))
        tbl_border = cast(TableBorder2, Props.get_property(prop_set=cell_rng, name="TableBorder2"))
        assert tbl_border.TopLine.Color == color.CommonColor.DARK_BLUE
        assert tbl_border.LeftLine.Color == color.CommonColor.DARK_BLUE
        assert tbl_border.RightLine.Color == color.CommonColor.BLACK
        assert tbl_border.BottomLine.Color == color.CommonColor.BLACK
        Lo.delay(delay)

        Calc.remove_border(sheet=sheet, cell_range=rng)
        Lo.delay(delay)

        # add_border(sheet: XSpreadsheet, range_name: str, color: int, border_vals: int)
        cell_rng = Calc.add_border(sheet=sheet, cell_range=rng, color=color.CommonColor.BLACK)  # reset colors
        cell_rng = Calc.add_border(
            sheet=sheet, range_name=rng_name, color=color.CommonColor.GREEN, border_vals=int(bval)
        )
        tbl_border = cast(TableBorder2, Props.get_property(prop_set=cell_rng, name="TableBorder2"))
        assert tbl_border.TopLine.Color == color.CommonColor.GREEN
        assert tbl_border.LeftLine.Color == color.CommonColor.GREEN
        assert tbl_border.RightLine.Color == color.CommonColor.BLACK
        assert tbl_border.BottomLine.Color == color.CommonColor.BLACK
        Lo.delay(delay)

        Calc.remove_border(sheet=sheet, cell_range=rng)
        Lo.delay(delay)

        cell_rng = Calc.add_border(sheet=sheet, cell_range=rng, color=color.CommonColor.BLACK)  # reset colors
        cell_rng = Calc.add_border(sheet, rng_name, color.CommonColor.DARK_BLUE, int(bval))
        tbl_border = cast(TableBorder2, Props.get_property(prop_set=cell_rng, name="TableBorder2"))
        assert tbl_border.TopLine.Color == color.CommonColor.DARK_BLUE
        assert tbl_border.LeftLine.Color == color.CommonColor.DARK_BLUE
        assert tbl_border.RightLine.Color == color.CommonColor.BLACK
        assert tbl_border.BottomLine.Color == color.CommonColor.BLACK
        Lo.delay(delay)

        Calc.remove_border(sheet=sheet, cell_range=rng)
        Lo.delay(delay)

        # add_border(sheet: XSpreadsheet, cell_range: XCellRange, color: int, border_vals: BorderEnum)
        cell_rng = Calc.add_border(sheet=sheet, cell_range=rng, color=color.CommonColor.BLACK)  # reset colors
        cell_rng = Calc.add_border(sheet=sheet, cell_range=rng, color=color.CommonColor.GREEN, border_vals=bval)
        tbl_border = cast(TableBorder2, Props.get_property(prop_set=cell_rng, name="TableBorder2"))
        assert tbl_border.TopLine.Color == color.CommonColor.GREEN
        assert tbl_border.LeftLine.Color == color.CommonColor.GREEN
        assert tbl_border.RightLine.Color == color.CommonColor.BLACK
        assert tbl_border.BottomLine.Color == color.CommonColor.BLACK
        Lo.delay(delay)

        Calc.remove_border(sheet=sheet, cell_range=rng)
        Lo.delay(delay)

        cell_rng = Calc.add_border(sheet=sheet, cell_range=rng, color=color.CommonColor.BLACK)  # reset colors
        cell_rng = Calc.add_border(sheet, rng, color.CommonColor.DARK_BLUE, bval)
        tbl_border = cast(TableBorder2, Props.get_property(prop_set=cell_rng, name="TableBorder2"))
        assert tbl_border.TopLine.Color == color.CommonColor.DARK_BLUE
        assert tbl_border.LeftLine.Color == color.CommonColor.DARK_BLUE
        assert tbl_border.RightLine.Color == color.CommonColor.BLACK
        assert tbl_border.BottomLine.Color == color.CommonColor.BLACK
        Lo.delay(delay)

        Calc.remove_border(sheet=sheet, cell_range=rng)
        Lo.delay(delay)

        # add_border(sheet: XSpreadsheet, range_name: str, color: int, border_vals: BorderEnum)
        cell_rng = Calc.add_border(sheet=sheet, cell_range=rng, color=color.CommonColor.BLACK)  # reset colors
        cell_rng = Calc.add_border(sheet=sheet, range_name=rng_name, color=color.CommonColor.GREEN, border_vals=bval)
        tbl_border = cast(TableBorder2, Props.get_property(prop_set=cell_rng, name="TableBorder2"))
        assert tbl_border.TopLine.Color == color.CommonColor.GREEN
        assert tbl_border.LeftLine.Color == color.CommonColor.GREEN
        assert tbl_border.RightLine.Color == color.CommonColor.BLACK
        assert tbl_border.BottomLine.Color == color.CommonColor.BLACK
        Lo.delay(delay)

        Calc.remove_border(sheet=sheet, cell_range=rng)
        Lo.delay(delay)

        cell_rng = Calc.add_border(sheet=sheet, cell_range=rng, color=color.CommonColor.BLACK)  # reset colors
        cell_rng = Calc.add_border(sheet, rng_name, color.CommonColor.DARK_BLUE, bval)
        tbl_border = cast(TableBorder2, Props.get_property(prop_set=cell_rng, name="TableBorder2"))
        assert tbl_border.TopLine.Color == color.CommonColor.DARK_BLUE
        assert tbl_border.LeftLine.Color == color.CommonColor.DARK_BLUE
        assert tbl_border.RightLine.Color == color.CommonColor.BLACK
        assert tbl_border.BottomLine.Color == color.CommonColor.BLACK
        Lo.delay(delay)

        with pytest.raises(TypeError):
            # error on unused key
            Calc.add_border(sheet=sheet, cellRange=rng, color=color.CommonColor.BLACK)
        with pytest.raises(TypeError):
            # error on incorrect number of args
            Calc.add_border(sheet=sheet)
    finally:
        Lo.close(closeable=doc, deliver_ownership=False)


def test_highlight_range(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc
    from ooodev.utils.gui import GUI
    from ooodev.utils.color import CommonColor

    visible = False
    delay = 0 # 3000
    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    try:
        sheet = Calc.get_sheet(doc=doc, index=0)
        if visible:
            GUI.set_visible(is_visible=visible, odoc=doc)

        rng_name = "B3:F8"
        headline = "Hello World!"
        rng = sheet.getCellRangeByName(rng_name)
        first = Calc.highlight_range(sheet=sheet, headline=headline, cell_range=rng)
        Lo.delay(delay)
        assert first is not None
        result = Calc.get_string(cell=first)
        assert result == headline
    finally:
        Lo.close(closeable=doc, deliver_ownership=False)
    Lo.delay(500)

    # highlight_range(sheet: XSpreadsheet,  headline: str, cell_range: XCellRange)
    doc = Calc.create_doc(loader)
    try:
        sheet = Calc.get_sheet(doc=doc, index=0)
        rng = sheet.getCellRangeByName(rng_name)
        if visible:
            GUI.set_visible(is_visible=visible, odoc=doc)
        first = Calc.highlight_range(sheet, headline, rng)
        Lo.delay(delay)
        assert first is not None
        result = Calc.get_string(cell=first)
        assert result == headline
    finally:
        Lo.close(closeable=doc, deliver_ownership=False)
    Lo.delay(500)
    
    doc = Calc.create_doc(loader)
    try:
        sheet = Calc.get_sheet(doc=doc, index=0)
        if visible:
            GUI.set_visible(is_visible=visible, odoc=doc)

        rng_name = "B3:F8"
        headline = "Hello World!"
        rng = sheet.getCellRangeByName(rng_name)
        first = Calc.highlight_range(sheet=sheet, headline=headline, cell_range=rng, color=CommonColor.LIGHT_GOLDENROD_YELLOW)
        Lo.delay(delay)
        assert first is not None
        result = Calc.get_string(cell=first)
        assert result == headline
    finally:
        Lo.close(closeable=doc, deliver_ownership=False)
    Lo.delay(500)


    # highlight_range(sheet: XSpreadsheet,  headline: str, range_name: str)
    doc = Calc.create_doc(loader)
    try:
        sheet = Calc.get_sheet(doc=doc, index=0)
        rng = sheet.getCellRangeByName(rng_name)
        if visible:
            GUI.set_visible(is_visible=visible, odoc=doc)
        first = Calc.highlight_range(sheet=sheet, headline=headline, range_name=rng_name)
        Lo.delay(delay)
        assert first is not None
        result = Calc.get_string(cell=first)
        assert result == headline
    finally:
        Lo.close(closeable=doc, deliver_ownership=False)
    Lo.delay(500)

    doc = Calc.create_doc(loader)
    try:
        sheet = Calc.get_sheet(doc=doc, index=0)
        rng = sheet.getCellRangeByName(rng_name)
        if visible:
            GUI.set_visible(is_visible=visible, odoc=doc)
        first = Calc.highlight_range(sheet, headline, rng_name)
        Lo.delay(delay)
        assert first is not None
        result = Calc.get_string(cell=first)
        assert result == headline

        with pytest.raises(TypeError):
            # error on unused key
            Calc.highlight_range(sheet=sheet, headline=headline, rangeName=rng_name)
        with pytest.raises(TypeError):
            # error on incorrect number of args
            Calc.highlight_range(sheet, headline)
    finally:
        Lo.close(closeable=doc, deliver_ownership=False)


def test_set_col_width(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.utils.props import Props
    from ooodev.office.calc import Calc

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    sheet = Calc.get_sheet(doc=doc, index=0)
    idx = 3
    width = 12
    cell_range = Calc.set_col_width(sheet=sheet, width=width, idx=idx)
    assert cell_range is not None
    # convert to decimal and use approx to test with a tolerance value.
    # https://docs.pytest.org/en/latest/reference/reference.html?highlight=approx#pytest.approx
    c_width = Props.get_property(prop_set=cell_range, name="Width")
    assert c_width / 10000 == pytest.approx(width / 100, rel=1e-2)
    Lo.close(closeable=doc, deliver_ownership=False)


def test_set_row_height(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.utils.props import Props
    from ooodev.office.calc import Calc

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    sheet = Calc.get_sheet(doc=doc, index=0)
    idx = 3
    height = 14
    cell_range = Calc.set_row_height(sheet=sheet, height=height, idx=idx)
    assert cell_range is not None
    c_height = Props.get_property(prop_set=cell_range, name="Height")
    assert c_height is not None
    # assert 0.1401 == 0.14 ± 1.4e-07
    # E         comparison failed
    # E         Obtained: 0.1401
    # E         Expected: 0.14 ± 1.4e-07
    # height can vary a small amount from set value.
    # convert to decimal and use approx to test with a tolerance value.
    # https://docs.pytest.org/en/latest/reference/reference.html?highlight=approx#pytest.approx
    assert c_height / 10000 == pytest.approx(height / 100, rel=1e-2)
    Lo.close(closeable=doc, deliver_ownership=False)


# endregion cell decoration

# region    scenarios
def test_scenarios(loader) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc
    from ooodev.utils.gui import GUI

    visible = False
    delay = 0
    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    sheet = Calc.get_sheet(doc=doc, index=0)
    if visible:
        GUI.set_visible(is_visible=visible, odoc=doc)
    vals = [[11, 12], ["Test13", "Test14"]]
    scenario1 = Calc.insert_scenario(
        sheet=sheet, range_name="B10:C11", vals=vals, name="First Scenario", comment="1st scenario."
    )
    Lo.delay(delay)
    assert scenario1 is not None
    assert scenario1.Name == "First Scenario"

    vals[0][0] = "Test21"
    vals[0][1] = "Test22"
    vals[1][0] = 23
    vals[1][1] = 24
    scenario2 = Calc.insert_scenario(
        sheet=sheet, range_name="B10:C11", vals=vals, name="Second Scenario", comment="Visible scenario."
    )
    Lo.delay(delay)
    assert scenario2 is not None
    assert scenario2.Name == "Second Scenario"

    vals[0][0] = 31
    vals[0][1] = 32
    vals[1][0] = "Test33"
    vals[1][1] = "Test34"
    scenario3 = Calc.insert_scenario(
        sheet=sheet, range_name="B10:C11", vals=vals, name="Third Scenario", comment="Last scenario."
    )
    Lo.delay(delay)
    assert scenario3 is not None
    assert scenario3.Name == "Third Scenario"

    scenario_apply = Calc.apply_scenario(sheet=sheet, name="Second Scenario")
    Lo.delay(delay)
    assert scenario_apply is not None
    assert scenario2.Name == "Second Scenario"

    Lo.close(closeable=doc, deliver_ownership=False)


# endregion scenarios
