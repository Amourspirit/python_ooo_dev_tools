from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])


def test_get_sheet(loader) -> None:
    # get_sheet is overload method.
    # testing each overload.
    from ooodev.calc import CalcDoc

    doc = CalcDoc.create_doc(loader)
    try:
        sheet_names = doc.get_sheet_names()
        assert len(sheet_names) == 1

        assert sheet_names[0] == "Sheet1"
        sheet = doc.get_sheet(sheet_name="Sheet1")

        sheet_name = sheet.get_sheet_name()
        assert sheet_name == "Sheet1"

        cell = sheet.get_cell(cell_name="A1")
        cell.set_val("test")
        val = cell.get_val()
        assert val == "test"
        assert cell.value == val

        text = cell.component.getText()
        val = text.getString()
        assert val == "test"

        cell = sheet.get_cell(cell_name="A2")
        cell.value = 2.5
        assert cell.value == 2.5

        assert doc.sheets[0].sheet_name == "Sheet1"
        assert doc.sheets["Sheet1"].sheet_name == "Sheet1"
        assert doc.sheets[-1].name == "Sheet1"
        doc.sheets[-1].name = "My Sheet"
        assert doc.sheets[-1].name == "My Sheet"

        doc.sheets.insert_new_by_name("Last Sheet", -1)

        assert doc.sheets.get_index_by_name("Last Sheet") == 1

    finally:
        doc.close()


def test_get_other_cells(loader) -> None:
    # get_sheet is overload method.
    # testing each overload.
    from ooodev.loader.lo import Lo
    from ooodev.calc import CalcDoc

    doc = CalcDoc.create_doc(loader)
    try:
        sheet_names = doc.get_sheet_names()
        assert len(sheet_names) == 1

        assert sheet_names[0] == "Sheet1"
        sheet = doc.get_sheet(sheet_name="Sheet1")

        sheet_name = sheet.get_sheet_name()
        assert sheet_name == "Sheet1"
        assert sheet.name == sheet_name

        cell = sheet.get_cell(cell_name="A1")
        cell.set_val("A1")
        val = cell.get_val()
        assert val == "A1"

        cell = sheet.get_cell("A1")
        cell.set_val("A1")
        val = cell.get_val()
        assert val == "A1"

        cell = sheet["A1"]
        val = cell.get_val()
        assert val == "A1"

        sheet["A1"].set_val("A ONE")
        assert sheet["A1"].get_val() == "A ONE"

        cell_b1 = cell.get_cell_right()
        assert cell_b1.cell_obj.col == "B"
        assert cell_b1.cell_obj.row == 1

        cell_a2 = cell.get_cell_down()
        assert cell_a2.cell_obj.col == "A"
        assert cell_a2.cell_obj.row == 2

        cell_b2 = cell_a2.get_cell_right()
        assert cell_b2.cell_obj.col == "B"
        assert cell_b2.cell_obj.row == 2

        cell_a1 = cell_b2.get_cell_up().get_cell_left()
        assert cell_a1.cell_obj.col == "A"
        assert cell_a1.cell_obj.row == 1

    finally:
        doc.close_doc()


# def test_sheet_row(loader) -> None:
#     from ooodev.loader.lo import Lo
#     from ooodev.office.calc import Calc
#     from ooodev.calc import CalcDoc

#     assert loader is not None
#     vals = get_vals()
#     doc = CalcDoc(Calc.create_doc(loader))
#     try:
#         sheet = doc.get_active_sheet()
#         sheet.set_array(values=vals, name="A1")

#         row = sheet.get_row_range(idx=0)

#         for cell in row:
#             assert cell is not None

#     finally:
#         Lo.close_doc(doc.component)


def test_insert_remove_sheet(loader) -> None:
    from ooodev.calc import Calc
    from ooodev.calc import CalcDoc

    doc = CalcDoc.create_doc(loader)
    try:
        sheet_names = doc.get_sheet_names()
        assert len(sheet_names) == 1
        assert sheet_names[0] == "Sheet1"

        assert len(doc.sheets) == 1

        sheet = doc.get_sheet(sheet_name="Sheet1")
        assert sheet.get_sheet_name() == "Sheet1"
        assert sheet.name == "Sheet1"
        doc.insert_sheet(name="Sheet2")
        assert len(doc.sheets) == 2

        sheet2 = doc.get_sheet(sheet_name="Sheet2")
        assert sheet2.get_sheet_name() == "Sheet2"
        assert sheet2.sheet_name == "Sheet2"
        assert sheet2.sheet_index == 1

        assert doc.remove_sheet(sheet2.sheet_name)
        sheet_names = doc.get_sheet_names()
        assert len(sheet_names) == 1
        assert sheet_names[0] == "Sheet1"

        for i in range(8):
            doc.sheets.insert_new_by_name(f"Sheet {i}", -1)
        assert len(doc.sheets) == 9

        del doc.sheets[0]
        assert len(doc.sheets) == 8

        names = doc.sheets.get_sheet_names()
        assert len(names) == 8
        assert names[-1] == "Sheet 7"
        del doc.sheets[-1]
        assert len(doc.sheets) == 7

        sheet = doc.sheets[-1]
        names = doc.sheets.get_sheet_names()
        assert len(names) == 7
        assert names[-1] == "Sheet 6"
        assert sheet.sheet_name == "Sheet 6"

        del doc.sheets[sheet.name]
        assert len(doc.sheets) == 6

        sheet = doc.sheets[-1]
        del doc.sheets[sheet]
        assert len(doc.sheets) == 5

    finally:
        doc.close_doc()


def test_calc_sheet(loader) -> None:
    from ooodev.calc import Calc
    from ooodev.calc import CalcDoc

    doc = CalcDoc.create_doc(loader)
    try:
        sheet_names = doc.sheets.get_sheet_names()
        assert len(sheet_names) == 1
        assert sheet_names[0] == "Sheet1"

        assert doc.sheets.has_by_name("Sheet1")

        sheet = doc.get_sheet("Sheet1")
        assert sheet.get_sheet_name() == "Sheet1"

        sheet = doc.get_sheet(sheet_name="Sheet1")
        assert sheet.get_sheet_name() == "Sheet1"

        assert doc.sheets.has_by_name("Sheet2") is False
        doc.sheets.insert_new_by_name("Sheet2", 1)

        for sheet in doc.sheets:
            assert sheet is not None
            assert sheet.sheet_name in ("Sheet1", "Sheet2")
            sheet["A1"].set_val("test")
            assert sheet["A1"].get_val() == "test"

        sheet2 = doc.get_sheet("Sheet2")
        assert sheet2.get_sheet_name() == "Sheet2"
        assert sheet2.sheet_name == "Sheet2"
        assert sheet2.sheet_index == 1

        doc.sheets.remove_by_name(sheet2.sheet_name)
        sheet_names = doc.sheets.get_sheet_names()
        assert len(sheet_names) == 1
        assert sheet_names[0] == "Sheet1"

        sheet = doc.sheets[0]
        assert sheet.sheet_name == "Sheet1"

        sheet = doc.sheets["Sheet1"]
        assert sheet.sheet_name == "Sheet1"

        assert doc.sheets["Sheet1"]["A1"].get_val() == "test"
        assert doc.sheets[0]["A1"].get_val() == "test"

        doc.sheets[0]["A2"].set_val("TEST2")
        assert doc.sheets["Sheet1"]["A2"].get_val() == "TEST2"
    finally:
        doc.close_doc()


def test_merge_unmerge(loader) -> None:
    # get_sheet is overload method.
    # testing each overload.
    from ooodev.loader.lo import Lo
    from ooodev.calc import Calc
    from ooodev.calc import CalcDoc
    from ooodev.utils.data_type.range_obj import RangeObj

    doc = Calc.create_doc(loader)
    try:
        calc_doc = CalcDoc(doc)
        sheet = calc_doc.get_active_sheet()
        range_obj = RangeObj.from_range("A1:B2")
        calc_rng = sheet.get_range(range_obj=range_obj)

        calc_rng.merge_cells(center=True)
        calc_rng.set_val("test")
        assert calc_rng.get_val() == "test"

        assert calc_rng.is_merged_cells()
        calc_rng.unmerge_cells()
        assert not calc_rng.is_merged_cells()
    finally:
        Lo.close_doc(doc)


def test_move_sheet(loader) -> None:
    from ooodev.loader.lo import Lo
    from ooodev.calc import Calc
    from ooodev.calc import CalcDoc
    from ooodev.utils.data_type.range_obj import RangeObj

    doc = Calc.create_doc(loader)
    try:
        calc_doc = CalcDoc(doc)
        my_sheet = calc_doc.insert_sheet(name="mysheet", idx=1)
        assert calc_doc.move_sheet(my_sheet.sheet_name, 0)
        sheet_names = calc_doc.get_sheet_names()
        assert len(sheet_names) == 2
        assert sheet_names[0] == "mysheet"
        assert sheet_names[1] == "Sheet1"

    finally:
        Lo.close_doc(doc)


def test_set_sheet_name(loader) -> None:
    from ooodev.loader.lo import Lo
    from ooodev.calc import Calc
    from ooodev.calc import CalcDoc

    doc = Calc.create_doc(loader)
    try:
        calc_doc = CalcDoc(doc)
        sheet_names = calc_doc.get_sheet_names()
        assert len(sheet_names) == 1
        assert sheet_names[0] == "Sheet1"
        sheet = calc_doc.get_sheet(sheet_name="Sheet1")
        sheet.set_sheet_name("mysheet")
        sheet_names = calc_doc.get_sheet_names()
        assert len(sheet_names) == 1
        assert sheet_names[0] == "mysheet"

    finally:
        Lo.close_doc(doc)


def test_get_view(loader) -> None:
    from ooodev.loader.lo import Lo
    from ooodev.calc import Calc
    from ooodev.calc import CalcDoc

    doc = Calc.create_doc(loader)
    try:
        calc_doc = CalcDoc(doc)
        view = calc_doc.get_view()
        assert view is not None
        sheet = view.component.getActiveSheet()
        name = Calc.get_sheet_name(sheet)
        assert name == "Sheet1"

    finally:
        Lo.close_doc(doc)


def test_get_set_active_sheet(loader) -> None:
    from ooodev.calc import CalcDoc

    doc = CalcDoc.create_doc(loader)
    try:
        sheet = doc.get_active_sheet()
        assert sheet is not None
        name = sheet.get_sheet_name()
        assert name == "Sheet1"
        new_sht = doc.insert_sheet(name="mysheet", idx=1)
        assert new_sht is not None
        lst_sheet = doc.insert_sheet(name="last", idx=1)
        assert lst_sheet is not None
        new_sht.set_active()

        active_sheet = doc.get_active_sheet()
        assert active_sheet is not None
        name = active_sheet.get_sheet_name()
        assert name == "mysheet"

    finally:
        doc.close()


def test_get_selected_cell(loader) -> None:
    from ooodev.loader.lo import Lo
    from ooodev.calc import CalcDoc

    assert loader is not None

    dispatched = False

    def on_dispatching(source, args):
        # This event will only be called when doc.dispatch_cmd() is called.
        # A content manager is used to subscribe and unsubscribe to the global events.
        nonlocal dispatched
        dispatched = True

    doc = CalcDoc.create_doc(loader)
    doc.subscribe_event("lo_dispatching", on_dispatching)
    try:
        sheet = doc.get_active_sheet()
        rng = sheet.get_range(range_name="B1:D4")
        cursor = rng.create_cursor()
        view = doc.get_view()
        result = view.select(cursor)
        assert result

        # uses CalcDoc.dispatch_cmd() to go to cell.
        # this will trigger the on_dispatching event.
        _ = sheet.goto_cell(cell_name="B4")
        assert dispatched is True
        selected_cell = sheet.get_selected_cell()
        assert selected_cell.cell_obj.col == "B"
        assert selected_cell.cell_obj.row == 4
    finally:
        Lo.close_doc(doc.component)


def test_insert_row(loader) -> None:
    from ooodev.loader.lo import Lo
    from ooodev.calc import CalcDoc

    assert loader is not None
    doc = CalcDoc.create_doc(loader)
    try:
        sheet = doc.get_active_sheet()
        cell = sheet["A2"]
        cell.set_val("test")
        assert sheet.insert_row(idx=1, count=2)

        cell = sheet.get_cell(cell_name="A4")
        assert cell.get_val() == "test"

    finally:
        Lo.close_doc(doc.component)


def test_delete_row(loader) -> None:
    from ooodev.loader.lo import Lo
    from ooodev.office.calc import Calc
    from ooodev.calc import CalcDoc

    assert loader is not None
    doc = CalcDoc.create_doc(loader)
    try:
        sheet = doc.get_active_sheet()
        cell_obj = Calc.get_cell_obj("A5")
        cell = sheet[cell_obj]
        cell.set_val("test")
        assert sheet.delete_row(idx=1, count=3)

        cell = sheet.get_cell("A2")
        assert cell.get_val() == "test"

    finally:
        Lo.close_doc(doc.component)


def test_insert_col(loader) -> None:
    from ooodev.loader.lo import Lo
    from ooodev.calc import CalcDoc

    assert loader is not None
    doc = CalcDoc.create_doc(loader)
    try:
        sheet = doc.get_active_sheet()
        cell = sheet.get_cell(cell_name="C2")
        cell.set_val("test")
        assert sheet.insert_column(idx=1, count=2)

        cell = sheet.get_cell(cell_name="E2")
        assert cell.get_val() == "test"

    finally:
        doc.close()


def test_del_col(loader) -> None:
    from ooodev.calc import CalcDoc

    assert loader is not None
    doc = CalcDoc.create_doc(loader)
    try:
        sheet = doc.get_active_sheet()
        cell = sheet.get_cell(cell_name="E2")
        cell.set_val("test")
        assert sheet.delete_column(idx=2, count=2)

        cell = sheet.get_cell(cell_name="C2")
        assert cell.get_val() == "test"

    finally:
        doc.close()


def test_insert_cells_down_rng(loader) -> None:
    from ooodev.loader.lo import Lo
    from ooodev.calc import CalcDoc

    assert loader is not None
    doc = CalcDoc.create_doc(loader)
    try:
        sheet = doc.get_active_sheet()
        rng = sheet.get_range(range_name="B1:D4")
        cell = sheet.get_cell(cell_name="B5")
        cell.set_val("test")
        assert sheet.insert_cells(rng.range_obj, False)

        cell = sheet.get_cell(cell_name="B9")
        assert cell.get_val() == "test"

    finally:
        doc.close()


def test_insert_cells_right(loader) -> None:
    from ooodev.loader.lo import Lo
    from ooodev.calc import CalcDoc

    assert loader is not None
    doc = CalcDoc.create_doc(loader)
    try:
        sheet = doc.get_active_sheet()
        rng = sheet.get_range(range_name="B1:D4")
        cell = sheet.get_cell(cell_name="D4")
        cell.set_val("test")
        assert sheet.insert_cells(rng.range_obj, True)

        cell = sheet.get_cell(cell_name="G4")
        assert cell.get_val() == "test"

    finally:
        doc.close()


def test_delete_cells_down(loader) -> None:
    from ooodev.loader.lo import Lo
    from ooodev.calc import CalcDoc

    assert loader is not None
    doc = CalcDoc.create_doc(loader)
    try:
        sheet = doc.get_active_sheet()
        rng = sheet.get_range(range_name="B1:D4")
        cell = sheet.get_cell(cell_name="B9")
        cell.set_val("test")
        assert sheet.delete_cells(rng.range_obj, False)

        cell = sheet.get_cell(cell_name="B5")
        assert cell.get_val() == "test"

    finally:
        doc.close()


def test_delete_cells_left(loader) -> None:
    from ooodev.loader.lo import Lo
    from ooodev.calc import CalcDoc

    assert loader is not None
    vals = get_vals()
    doc = CalcDoc.create_doc(loader)
    try:
        sheet = doc.get_active_sheet()
        sheet.set_array(values=vals, name="A1")

        rng = sheet.find_used()
        cell = sheet.get_cell(cell_name="D5")
        cell.set_val("test")
        assert sheet.delete_cells(rng.range_obj, True)

        cell = sheet.get_cell(cell_name="A5")
        assert cell.get_val() == "test"

    finally:
        doc.close()


def test_clear_cells(loader) -> None:
    from ooodev.office.calc import CellFlagsEnum
    from ooodev.calc import CalcDoc

    assert loader is not None
    vals = get_vals()
    doc = CalcDoc.create_doc(loader)

    try:
        flags = CellFlagsEnum.VALUE | CellFlagsEnum.STRING
        sheet = doc.get_active_sheet()
        sheet.set_array(values=vals, name="A1")

        rng = sheet.find_used()
        sheet.clear_cells(range_val=rng.range_obj, cell_flags=flags)
        cell = sheet.get_cell(cell_name="A1")
        assert cell.get_val() is None
    finally:
        doc.close()


def test_clear_cells_range(loader) -> None:
    from ooodev.calc import CalcDoc

    assert loader is not None
    vals = get_vals()
    doc = CalcDoc.create_doc(loader)
    try:
        sheet = doc.get_active_sheet()
        sheet.set_array(values=vals, name="A1")

        rng = sheet.find_used()
        assert rng.clear_cells()

        cell = sheet.get_cell(cell_name="A1")
        assert cell.get_val() is None

    finally:
        doc.close()


def test_get_number(loader) -> None:
    from ooodev.calc import CalcDoc

    vals = get_vals()
    doc = CalcDoc.create_doc(loader)
    try:
        sheet = doc.get_active_sheet()
        sheet.set_array(values=vals, name="A1")

        cell = sheet.get_cell(cell_name="C2")
        assert cell.get_num() == 3.0
        cell = sheet.get_cell(cell_name="C3")
        assert cell.get_num() == 7.0

    finally:
        doc.close()


def test_set_arr_by_range(loader) -> None:
    from ooodev.calc import CalcDoc
    from ooodev.utils.table_helper import TableHelper

    doc = CalcDoc.create_doc(loader)
    try:
        arr1 = ((1, 2, 3), (4, 5, 6), (7, 8, 9))
        arr = TableHelper.to_2d_list(arr1)
        sheet = doc.get_active_sheet()
        rng = sheet.get_range(range_name="A1:C3")
        rng.set_array(arr)

        cell = sheet.get_cell(cell_name="A1")
        assert cell.get_num() == 1.0
        cell = sheet.get_cell(cell_name="C2")
        assert cell.get_num() == 6.0
        cell = sheet.get_cell(cell_name="C3")
        assert cell.get_num() == 9.0
    finally:
        doc.close()


def test_get_array(loader) -> None:
    from ooodev.calc import CalcDoc

    vals = get_vals()
    doc = CalcDoc.create_doc(loader)
    try:
        sheet = doc.get_active_sheet()
        sheet.set_array(values=vals, name="A1")
        rng = sheet.find_used()

        arr1 = sheet.get_array(range_obj=rng.range_obj)
        arr2 = rng.get_array()

        for i in range(len(arr1)):
            for j in range(len(arr1[i])):
                assert arr1[i][j] == arr2[i][j]

    finally:
        doc.close()


def test_float_arr(loader) -> None:
    from ooodev.calc import CalcDoc
    from ooodev.utils.table_helper import TableHelper

    def arr_cb(row: int, col: int, prev_val) -> str:
        if row == 0 and col == 0:
            return "1"
        return str(int(prev_val) + 1)

    doc = CalcDoc.create_doc(loader)
    try:
        arr_size = 8
        arr = TableHelper.to_2d_tuple(TableHelper.make_2d_array(arr_size, arr_size, arr_cb))
        sheet = doc.get_active_sheet()
        sheet.set_array(values=arr, name="A1")
        rng = sheet.find_used()

        arr_float = rng.get_float_array()
        for row, row_data in enumerate(arr):
            for col, _ in enumerate(row_data):
                assert arr_float[row][col] == float(arr[row][col])

    finally:
        doc.close()


def test_set_date(loader) -> None:
    from ooodev.calc import CalcDoc

    doc = CalcDoc.create_doc(loader)
    try:
        sheet = doc.get_active_sheet()

        cell = sheet.get_cell(cell_name="A1")
        cell.set_date(day=22, month=11, year=2022)
        assert cell.get_string() == "44887.0"

    finally:
        doc.close()


def test_annotation(loader) -> None:
    from ooodev.calc import CalcDoc

    doc = CalcDoc.create_doc(loader)
    try:
        sheet = doc.get_active_sheet()

        cn = "B2"
        msg = "Hello World"
        cell = sheet.get_cell(cell_name=cn)
        ann = cell.add_annotation(msg)
        assert ann is not None
        ann_str = cell.get_annotation_str()
        assert ann_str == msg

    finally:
        doc.close()


def test_is_single_cell_range(loader) -> None:
    from ooodev.calc import CalcDoc

    doc = CalcDoc.create_doc(loader)
    try:
        sheet = doc.get_active_sheet()
        rng = sheet.get_range(range_name="A1:A1")
        assert rng.is_single_cell_range()
        assert rng.is_single_column_range()
        assert rng.is_single_row_range()

        rng = sheet.get_range(range_name="A1:B1")
        assert not rng.is_single_cell_range()
        assert not rng.is_single_column_range()
        assert rng.is_single_row_range()

        rng = sheet.get_range(range_name="A1:A2")
        assert not rng.is_single_cell_range()
        assert rng.is_single_column_range()
        assert not rng.is_single_row_range()

    finally:
        doc.close()


def get_vals():
    return (
        ("Name", "Fruit", "Quantity"),
        ("Alice", "Apples", 3),
        ("Alice", "Oranges", 7),
        ("Bob", "Apples", 3),
        ("Alice", "Apples", 9),
        ("Bob", "Apples", 5),
        ("Bob", "Oranges", 6),
        ("Alice", "Oranges", 3),
        ("Alice", "Apples", 8),
        ("Alice", "Oranges", 1),
        ("Bob", "Oranges", 2),
        ("Bob", "Oranges", 7),
        ("Bob", "Apples", 1),
        ("Alice", "Apples", 8),
        ("Alice", "Oranges", 8),
        ("Alice", "Apples", 7),
        ("Bob", "Apples", 1),
        ("Bob", "Oranges", 9),
        ("Bob", "Oranges", 3),
        ("Alice", "Oranges", 4),
        ("Alice", "Apples", 9),
    )
