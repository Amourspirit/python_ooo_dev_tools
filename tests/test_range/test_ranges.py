import pytest

if __name__ == "__main__":
    pytest.main([__file__])


def test_get_range_values(loader) -> None:
    from ooodev.utils.data_type.range_values import RangeValues
    from ooodev.loader.lo import Lo
    from ooodev.office.calc import Calc

    doc = Calc.create_doc()

    try:

        rng_name = "A2:D6"
        rv1 = RangeValues.from_range(range_val=rng_name)
        assert rv1.col_start == 0
        assert rv1.row_start == 1
        assert rv1.col_end == 3
        assert rv1.row_end == 5
        assert str(rv1) == rng_name
        assert rv1 == rng_name

        rv2 = RangeValues.from_range(rng_name)
        assert str(rv2) == rng_name
        assert rv1 == rv2
        assert rv1 != "Roses are red"

        ro = rv2.get_range_obj()
        assert str(ro) == rng_name
        assert ro == rv1
        assert rv2 == ro
        assert ro == rng_name
        assert ro != "A1:C3"
        assert ro != "Roses are red"
        assert ro.sheet_name.startswith("Sheet")

        assert ro.start.col_obj.index == 0
        assert ro.start.col_obj.value == "A"
        assert ro.start.row_obj.index == 1
        assert ro.start.row_obj.value == 2

        assert ro.end.col_obj.index == 3
        assert ro.end.col_obj.value == "D"
        assert ro.end.row_obj.index == 5
        assert ro.end.row_obj.value == 6

        assert ro.col_start == "A"
        assert ro.col_end == "D"
        assert ro.row_start == 2
        assert ro.row_end == 6

        assert ro.start.col_obj < ro.end.col_obj
        assert ro.start.row_obj < ro.end.row_obj

        assert ro.end.col_obj > ro.start.col_obj
        assert ro.end.row_obj > ro.start.row_obj

        assert ro.start.col_obj != ro.end.col_obj
        assert ro.start.row_obj != ro.end.row_obj
    finally:
        Lo.close_doc(doc)


@pytest.mark.parametrize(
    ("val", "add", "end", "expected"),
    [
        (10, 10, True, 20),
        (0, 5, True, 5),
        (11, -5, True, 6),
        (10, 10, False, 0),
        (6, 5, False, 1),
        (11, -5, False, 16),
    ],
)
def test_range_values_add_cols(val: int, add: int, end: bool, expected: int) -> None:
    from ooodev.utils.data_type.range_values import RangeValues

    col_start = 0
    col_end = 0
    if end:
        col_end = val
    else:
        col_start = val
    rv1 = RangeValues(col_start=col_start, col_end=col_end, row_start=0, row_end=0, sheet_idx=0)
    rv2 = rv1.add_cols(add, end)
    if end:
        assert rv2.col_end == expected
    else:
        assert rv2.col_start == expected


@pytest.mark.parametrize(
    ("start", "end", "add", "to_end", "expected_start", "expected_end"),
    [
        (0, 10, 10, True, 0, 20),
        (0, 0, 5, True, 0, 5),
        (0, 11, -5, True, 0, 6),
        (10, 11, 10, False, 0, 11),
        (6, 8, 5, False, 1, 8),
        (6, 11, -5, False, 11, 11),
    ],
)
def test_range_values_add_rows(
    start: int, end: int, add: int, to_end: bool, expected_start: int, expected_end: int
) -> None:
    from ooodev.utils.data_type.range_values import RangeValues

    rv1 = RangeValues(col_start=0, col_end=0, row_start=start, row_end=end, sheet_idx=0)
    rv2 = rv1.add_rows(add, to_end)
    assert rv2.row_end == expected_end
    assert rv2.row_start == expected_start


@pytest.mark.parametrize(
    ("start", "end", "add", "to_end", "expected_start", "expected_end"),
    [
        (0, 10, 10, True, 0, 20),
        (0, 0, 5, True, 0, 5),
        (0, 11, -5, True, 0, 6),
        (10, 10, 10, False, 0, 10),
        (5, 6, 5, False, 0, 6),
        (0, 11, -5, False, 5, 11),
    ],
)
def test_range_values_add_cols(
    start: int, end: int, add: int, to_end: bool, expected_start: int, expected_end: int
) -> None:
    from ooodev.utils.data_type.range_values import RangeValues

    # RangeValues will automatically swap columns when start is greater the end
    rv1 = RangeValues(col_start=start, col_end=end, row_start=0, row_end=0, sheet_idx=0)
    rv2 = rv1.add_cols(add, to_end)
    assert rv2.col_start == expected_start
    assert rv2.col_end == expected_end


@pytest.mark.parametrize(
    ("start", "end", "subt", "to_end", "expected_start", "expected_end"),
    [
        (0, 10, 10, True, 0, 0),
        (0, 5, 0, True, 0, 5),
        (2, 7, 3, True, 2, 4),
        (0, 11, -5, True, 0, 16),
        (1, 12, 10, False, 11, 12),
        (6, 6, 5, False, 6, 11),
        (11, 13, -5, False, 6, 13),
    ],
)
def test_range_values_subtract_rows(
    start: int, end: int, subt: int, to_end: bool, expected_start: int, expected_end: int
) -> None:
    from ooodev.utils.data_type.range_values import RangeValues

    rv1 = RangeValues(col_start=0, col_end=0, row_start=start, row_end=end, sheet_idx=0)
    rv2 = rv1.subtract_rows(subt, to_end)
    assert rv2.row_start == expected_start
    assert rv2.row_end == expected_end


@pytest.mark.parametrize(
    ("start", "end", "subt", "to_end", "expected_start", "expected_end"),
    [
        (10, 10, 10, True, 0, 10),
        (1, 5, 0, True, 1, 5),
        (4, 7, 3, True, 4, 4),
        (1, 11, -5, True, 1, 16),
        (12, 14, 10, False, 14, 22),
        (6, 12, 5, False, 11, 12),
        (8, 11, -5, False, 3, 11),
    ],
)
def test_range_values_subtract_cols(
    start: int, end: int, subt: int, to_end: bool, expected_start: int, expected_end: int
) -> None:
    from ooodev.utils.data_type.range_values import RangeValues

    rv1 = RangeValues(col_start=start, col_end=end, row_start=0, row_end=0, sheet_idx=0)
    rv2 = rv1.subtract_cols(subt, to_end)
    assert rv2.col_start == expected_start
    assert rv2.col_end == expected_end


def test_range_math_errors() -> None:
    from ooodev.utils.data_type.range_values import RangeValues

    rv = RangeValues(col_start=1, col_end=1, row_start=1, row_end=1, sheet_idx=0)
    with pytest.raises(ValueError):
        _ = rv.add_rows(-2)

    with pytest.raises(ValueError):
        _ = rv.add_rows(2, False)

    with pytest.raises(ValueError):
        _ = rv.add_cols(-2)

    with pytest.raises(ValueError):
        _ = rv.add_cols(2, False)

    with pytest.raises(ValueError):
        _ = rv.subtract_cols(2)

    with pytest.raises(ValueError):
        _ = rv.subtract_cols(-2, False)

    with pytest.raises(ValueError):
        _ = rv.subtract_rows(2)

    with pytest.raises(ValueError):
        _ = rv.subtract_rows(-2, False)


def test_get_range_obj(loader) -> None:
    from ooodev.utils.data_type.range_obj import RangeObj
    from ooodev.loader.lo import Lo
    from ooodev.office.calc import Calc

    doc = Calc.create_doc()
    try:

        ro1 = RangeObj.from_range("A2:D6")
        assert ro1.col_start == "A"
        assert ro1.row_start == 2
        assert ro1.col_end == "D"
        assert ro1.row_end == 6
        assert ro1.sheet_name.startswith("Sheet")

        # test that RangObj assigns itself to CellObj.range_obj
        assert ro1.cell_start.range_obj is ro1
        assert ro1.cell_end.range_obj is ro1

        assert ro1.cell_start.col_obj.cell_obj is ro1.cell_start
        assert ro1.cell_start.row_obj.cell_obj is ro1.cell_start

        assert ro1.cell_end.col_obj.cell_obj is ro1.cell_end
        assert ro1.cell_end.row_obj.cell_obj is ro1.cell_end

        assert ro1.cell_start.col == "A"
        assert ro1.cell_start.row == 2

        assert ro1.cell_end.col == "D"
        assert ro1.cell_end.row == 6

        assert ro1.cell_start.col_obj.value == "A"
        assert ro1.cell_start.col_obj.index == 0

        assert ro1.cell_start.row_obj.value == 2
        assert ro1.cell_start.row_obj.index == 1

        assert ro1.cell_end.col_obj.value == "D"
        assert ro1.cell_end.col_obj.index == 3

        assert ro1.cell_end.row_obj.value == 6
        assert ro1.cell_end.row_obj.index == 5

        assert ro1.cell_start.sheet_idx == ro1.sheet_idx
        assert ro1.cell_end.sheet_idx == ro1.sheet_idx

        ro2 = RangeObj.from_range(range_val="a2:d6")
        assert ro2.col_start == "A"
        assert ro2.row_start == 2
        assert ro2.col_end == "D"
        assert ro2.row_end == 6

        assert ro1 == ro2
        assert ro1.start.col_obj.index == 0
        assert ro1.end.col_obj.index == 3

        ro2 = RangeObj.from_range(range_val="a2:d6")
        assert ro2 == ro1

        rv1 = ro1.get_range_values()
        assert str(rv1) == "A2:D6"
        assert rv1 == "A2:D6"
        assert rv1 != "A2:D7"
        assert rv1 != "Roses are red"

        rv2 = ro2.get_range_values()
        assert rv1 == rv2
        assert rv2 == rv1
        assert rv2 == ro2
        assert ro2 == rv2
        assert rv2 == "A2:D6"

        ro3 = RangeObj.from_range(rv2)
        assert ro3 == ro2
    finally:
        Lo.close_doc(doc)


def test_get_range_values_cell_address(loader) -> None:
    from ooodev.utils.data_type.range_values import RangeValues
    from ooodev.loader.lo import Lo
    from ooodev.office.calc import Calc

    doc = Calc.create_doc()

    try:

        rng_name = "A2:D6"
        rv1 = RangeValues.from_range(range_val=rng_name)
        assert rv1.col_start == 0
        assert rv1.row_start == 1
        assert rv1.col_end == 3
        assert rv1.row_end == 5
        assert str(rv1) == rng_name
        assert rv1 == rng_name
        cr1 = rv1.get_cell_range_address()
        assert cr1.StartColumn == rv1.col_start
        assert cr1.StartRow == rv1.row_start
        assert cr1.EndColumn == rv1.col_end
        assert cr1.EndRow == rv1.row_end

        rv2 = RangeValues.from_range(cr1)
        assert rv2 == rv1

    finally:
        Lo.close_doc(doc)


def test_get_range_obj_cell_address(loader) -> None:
    from ooodev.utils.data_type.range_obj import RangeObj

    from ooodev.loader.lo import Lo
    from ooodev.office.calc import Calc

    doc = Calc.create_doc()
    sheet = Calc.get_sheet(doc)

    try:

        rng_name = "A2:D6"
        rv1 = RangeObj.from_range(range_val=rng_name)
        cr1 = rv1.get_cell_range_address()
        addr = Calc.get_address(sheet=sheet, range_name=rng_name)
        assert Calc.is_equal_addresses(cr1, addr)

        rv2 = RangeObj.from_range(cr1)
        assert rv2 == rv1

    finally:
        Lo.close_doc(doc)


def test_range_obj_is_methods(loader) -> None:
    from ooodev.utils.data_type.range_obj import RangeObj
    from ooodev.utils.data_type.cell_obj import CellObj

    from ooodev.loader.lo import Lo
    from ooodev.office.calc import Calc

    doc = Calc.create_doc()

    try:

        rng_name = "A1:AA1"
        rng = RangeObj.from_range(range_val=rng_name)
        assert rng.is_single_cell() == False
        assert rng.is_single_col() == False
        assert rng.is_single_row() == True

        rng_name = "A1:A77"
        rng = RangeObj.from_range(range_val=rng_name)
        assert rng.is_single_cell() == False
        assert rng.is_single_col() == True
        assert rng.is_single_row() == False

        rng_name = "A1:R17"
        rng = RangeObj.from_range(range_val=rng_name)
        sub_rng = rng.get_start_col()
        assert sub_rng.is_single_cell() == False
        assert sub_rng.is_single_col() == True
        assert sub_rng.is_single_row() == False
        assert sub_rng.col_start == "A"
        assert sub_rng.col_end == "A"
        assert sub_rng.row_start == 1
        assert sub_rng.row_end == 17

        sub_rng = rng.get_end_col()
        assert sub_rng.is_single_cell() == False
        assert sub_rng.is_single_col() == True
        assert sub_rng.is_single_row() == False
        assert sub_rng.col_start == "R"
        assert sub_rng.col_end == "R"
        assert sub_rng.row_start == 1
        assert sub_rng.row_end == 17

        sub_rng = rng.get_start_row()
        assert sub_rng.is_single_cell() == False
        assert sub_rng.is_single_col() == False
        assert sub_rng.is_single_row() == True
        assert sub_rng.col_start == "A"
        assert sub_rng.col_end == "R"
        assert sub_rng.row_start == 1
        assert sub_rng.row_end == 1

        sub_rng = rng.get_end_row()
        assert sub_rng.is_single_cell() == False
        assert sub_rng.is_single_col() == False
        assert sub_rng.is_single_row() == True
        assert sub_rng.col_start == "A"
        assert sub_rng.col_end == "R"
        assert sub_rng.row_start == 17
        assert sub_rng.row_end == 17

        cell = CellObj(col="C", row=10, sheet_idx=0)
        rng = cell.get_range_obj()
        assert rng.is_single_cell() == True
        assert rng.is_single_col() == True
        assert rng.is_single_row() == True

    finally:
        Lo.close_doc(doc)


@pytest.mark.parametrize(
    ("rng_name", "cell_name", "expected"),
    [
        ("A2:B7", "A2", True),
        ("A2:B7", "B7", True),
        ("a2:b7", "A3", True),
        ("A2:B7", "a5", True),
        ("A2:B7", "A7", True),
        ("Sheet1.A2:B7", "Sheet1.A7", True),
        ("Sheet2.A2:B7", "Sheet2.A7", True),
        ("Sheet1.A2:B7", "Sheet2.A3", False),
        ("AA2:AB7", "AA3", True),
        ("AB2:AB7", "AA3", False),
        ("A2:B7", "A1", False),
        ("A2:B7", "A8", False),
        ("A2:B7", "C100", False),
        ("A2:B7", "B8", False),
    ],
)
def test_range_obj_contains(rng_name: str, cell_name: str, expected: bool, loader) -> None:
    from ooodev.loader.lo import Lo
    from ooodev.office.calc import Calc
    from ooodev.utils.data_type.cell_obj import CellObj
    from ooodev.utils.data_type.range_obj import RangeObj

    doc = Calc.create_doc()
    _ = Calc.insert_sheet(doc, "Sheet2", 1)
    _ = Calc.get_sheet(doc, 0)

    try:
        rng = RangeObj.from_range(rng_name)
        cell = CellObj.from_cell(cell_name)
        assert rng.contains(cell) == expected
    finally:
        Lo.close_doc(doc)


@pytest.mark.parametrize(
    ("rng_name", "rows", "cols", "count"),
    [
        ("A2:B7", 6, 2, 12),
        ("A2:A2", 1, 1, 1),
        ("A3:G25", 23, 7, 161),
        ("G16:K25", 10, 5, 50),
    ],
)
def test_row_col_cell_count(rng_name: str, rows: int, cols: int, count: int) -> None:
    from ooodev.utils.data_type.range_obj import RangeObj

    rng = RangeObj.from_range(rng_name)
    assert rng.row_count == rows
    assert rng.col_count == cols
    assert rng.cell_count == count


@pytest.mark.parametrize(
    ("rng_name",),
    [
        ("A1:C3",),
        ("A2:A2",),
        ("R22:AA55",),
        ("O30:AQ66",),
    ],
)
def test_gen(rng_name: str) -> None:
    from ooodev.utils.data_type.range_obj import RangeObj
    from ooodev.utils.table_helper import TableHelper

    parts = TableHelper.get_range_parts(rng_name)

    rng = RangeObj(
        col_start=parts.col_start, col_end=parts.col_end, row_start=parts.row_start, row_end=parts.row_end, sheet_idx=0
    )
    row_count = 0
    cell_count = 0
    for row, cells in enumerate(rng.get_cells()):
        row_count += 1
        for col, cell in enumerate(cells):
            cell_count += 1
            assert cell.col_obj.index == col + rng.start_col_index
            assert cell.row == row + rng.start_row_index + 1
    assert row_count == rng.row_count
    assert cell_count == rng.cell_count


# region Test RangeObj add/subtract rows using int


@pytest.mark.parametrize(
    ("rng_name", "rows", "expected"),
    [
        ("A1:C3", 1, "A1:C4"),
        ("A2:A2", 2, "A2:A4"),
        ("C2:R22", 10, "C2:R32"),
        ("C2:R22", -10, "C2:R12"),
    ],
)
def test_add_range_rows_end(rng_name: str, rows: int, expected: str) -> None:
    from ooodev.utils.data_type.range_obj import RangeObj

    # adding rows to the end is appending rows to the bottom of the range.
    # the overall rows is increased, unless adding negative
    # the first row is expected to remain the same
    # the columns are expected to remain the same

    rng1 = RangeObj.from_range(rng_name)
    rng2 = rng1 + rows
    assert str(rng2) == expected


@pytest.mark.parametrize(
    ("rng_name", "rows", "expected"),
    [
        ("A3:C3", 1, "A2:C3"),
        ("A3:A3", 2, "A1:A3"),
        ("C22:R30", 10, "C12:R30"),
        ("C12:R30", -10, "C22:R30"),
    ],
)
def test_add_range_rows_start(rng_name: str, rows: int, expected: str) -> None:
    from ooodev.utils.data_type.range_obj import RangeObj

    # adding rows to the start is adding rows to the top of the range.
    # The overall rows are increased by adding to the top of the range, unless adding negative.
    # the last row is expected to remain the same
    # the columns are expected to remain the same

    rng1 = RangeObj.from_range(rng_name)
    rng2 = rows + rng1
    assert str(rng2) == expected


@pytest.mark.parametrize(
    ("rng_name", "rows", "expected"),
    [
        ("A3:C4", 1, "A3:C3"),
        ("B33:C44", 10, "B33:C34"),
        ("B33:C44", -10, "B33:C54"),
    ],
)
def test_subtract_range_rows_end(rng_name: str, rows: int, expected: str) -> None:
    from ooodev.utils.data_type.range_obj import RangeObj

    # subtracting rows from end is removing rows from the bottom of the range.
    # the overall rows are reduced by removing from the bottom of the range, unless subtracting a negative value.
    # the columns are expected to remain the same

    rng1 = RangeObj.from_range(rng_name)
    rng2 = rng1 - rows
    assert str(rng2) == expected


@pytest.mark.parametrize(
    ("rng_name", "rows", "expected"),
    [
        ("A3:C4", 1, "A4:C4"),
        ("B33:C44", 10, "B43:C44"),
        ("B33:C44", -10, "B23:C44"),
    ],
)
def test_subtract_range_rows_start(rng_name: str, rows: int, expected: str) -> None:
    from ooodev.utils.data_type.range_obj import RangeObj

    # subtracting rows from start is removing rows from the top of the range.
    # the overall rows are reduced by removing from the top of the range, unless subtracting a negative value.
    # the columns are expected to remain the same

    rng1 = RangeObj.from_range(rng_name)
    rng2 = rows - rng1
    assert str(rng2) == expected


# endregion Test RangeObj add/subtract rows using int

# region Test RangeObj add/subtract cols using str


@pytest.mark.parametrize(
    ("rng_name", "col", "expected"),
    [
        ("A1:C3", "A", "A1:D3"),
        ("A2:A2", "B", "A2:C2"),
        ("C2:R22", "J", "C2:AB22"),
    ],
)
def test_add_range_cols_end(rng_name: str, col: str, expected: str) -> None:
    from ooodev.utils.data_type.range_obj import RangeObj

    # adding cols to the end is adding cols to the right of the range.
    # The overall cols are increased by adding to the right of the range.
    # the rows are expected to remain the same.
    # the first column is expect to remain the same.

    rng1 = RangeObj.from_range(rng_name)
    rng2 = rng1 + col
    assert str(rng2) == expected


@pytest.mark.parametrize(
    ("rng_name", "col", "expected"),
    [
        ("B1:C3", "A", "A1:C3"),
        ("D2:E2", "B", "B2:E2"),
        ("R2:V22", "J", "H2:V22"),
    ],
)
def test_add_range_cols_start(rng_name: str, col: str, expected: str) -> None:
    from ooodev.utils.data_type.range_obj import RangeObj

    # adding cols to the start is adding cols to the left of the range.
    # The overall cols are increased by adding to the left of the range.
    # the rows are expected to remain the same.
    # the last column is expect to remain the same.

    rng1 = RangeObj.from_range(rng_name)
    rng2 = col + rng1
    assert str(rng2) == expected


@pytest.mark.parametrize(
    ("rng_name", "col", "expected"),
    [
        ("A1:C3", "A", "A1:B3"),
        ("A2:D2", "B", "A2:B2"),
        ("C2:R22", "J", "C2:H22"),
    ],
)
def test_subtract_range_cols_end(rng_name: str, col: str, expected: str) -> None:
    from ooodev.utils.data_type.range_obj import RangeObj

    # subtracting cols from the end is removing cols from the right of the range.
    # The overall cols are decreased by removeing from the right of the range.
    # the rows are expected to remain the same.
    # the first column is expect to remain the same.

    rng1 = RangeObj.from_range(rng_name)
    rng2 = rng1 - col
    assert str(rng2) == expected


@pytest.mark.parametrize(
    ("rng_name", "col", "expected"),
    [
        ("A1:C3", "A", "B1:C3"),
        ("A2:D2", "B", "C2:D2"),
        ("C2:R22", "J", "M2:R22"),
    ],
)
def test_subtract_range_cols_start(rng_name: str, col: str, expected: str) -> None:
    from ooodev.utils.data_type.range_obj import RangeObj

    # subtracting cols from the start is removing cols from the left of the range.
    # The overall cols are decreased by removeing from the left of the range.
    # the rows are expected to remain the same.
    # the last column is expect to remain the same.

    rng1 = RangeObj.from_range(rng_name)
    rng2 = col - rng1
    assert str(rng2) == expected


# endregion Test RangeObj add/subtract cols using str


# region Test RangeObj add/subtract rows using RowObj
@pytest.mark.parametrize(
    ("rng_name", "rows", "expected"),
    [
        ("A1:C3", 1, "A1:C4"),
        ("A2:A2", 2, "A2:A4"),
        ("C2:R22", 10, "C2:R32"),
    ],
)
def test_add_range_rows_end_row_obj(rng_name: str, rows: int, expected: str) -> None:
    from ooodev.utils.data_type.range_obj import RangeObj
    from ooodev.utils.data_type.row_obj import RowObj

    # adding rows to the end is appending rows to the bottom of the range.
    # the overall rows is increased
    # the first row is expected to remain the same
    # the columns are expected to remain the same

    ro = RowObj(rows)

    rng1 = RangeObj.from_range(rng_name)
    rng2 = rng1 + ro
    assert str(rng2) == expected


@pytest.mark.parametrize(
    ("rng_name", "rows", "expected"),
    [
        ("A3:C3", 1, "A2:C3"),
        ("A3:A3", 2, "A1:A3"),
        ("C22:R30", 10, "C12:R30"),
    ],
)
def test_add_range_rows_start_row_obj(rng_name: str, rows: int, expected: str) -> None:
    from ooodev.utils.data_type.range_obj import RangeObj
    from ooodev.utils.data_type.row_obj import RowObj

    # adding rows to the start is adding rows to the top of the range.
    # The overall rows are increased by adding to the top of the range
    # the last row is expected to remain the same
    # the columns are expected to remain the same

    ro = RowObj(rows)
    rng1 = RangeObj.from_range(rng_name)
    rng2 = ro + rng1
    assert str(rng2) == expected


@pytest.mark.parametrize(
    ("rng_name", "rows", "expected"),
    [
        ("A3:C4", 1, "A3:C3"),
        ("B33:C44", 10, "B33:C34"),
    ],
)
def test_subtract_range_rows_end_row_obj(rng_name: str, rows: int, expected: str) -> None:
    from ooodev.utils.data_type.range_obj import RangeObj
    from ooodev.utils.data_type.row_obj import RowObj

    # subtracting rows from end is removing rows from the bottom of the range.
    # the overall rows are reduced by removing from the bottom of the range.
    # the columns are expected to remain the same

    ro = RowObj(rows)
    rng1 = RangeObj.from_range(rng_name)
    rng2 = rng1 - ro
    assert str(rng2) == expected


@pytest.mark.parametrize(
    ("rng_name", "rows", "expected"),
    [
        ("A3:C4", 1, "A4:C4"),
        ("B33:C44", 10, "B43:C44"),
    ],
)
def test_subtract_range_rows_start_row_obj(rng_name: str, rows: int, expected: str) -> None:
    from ooodev.utils.data_type.range_obj import RangeObj
    from ooodev.utils.data_type.row_obj import RowObj

    # subtracting rows from start is removing rows from the top of the range.
    # the overall rows are reduced by removing from the top of the range.
    # the columns are expected to remain the same

    ro = RowObj(rows)
    rng1 = RangeObj.from_range(rng_name)
    rng2 = ro - rng1
    assert str(rng2) == expected


# endregion Test RangeObj add/subtract rows using RowObj


# region Test RangeObj add/subtract cols using ColObj
@pytest.mark.parametrize(
    ("rng_name", "col", "expected"),
    [
        ("A1:C3", "A", "A1:D3"),
        ("A2:A2", "B", "A2:C2"),
        ("C2:R22", "J", "C2:AB22"),
    ],
)
def test_add_range_cols_end_col_obj(rng_name: str, col: str, expected: str) -> None:
    from ooodev.utils.data_type.range_obj import RangeObj
    from ooodev.utils.data_type.col_obj import ColObj

    # adding cols to the end is adding cols to the right of the range.
    # The overall cols are increased by adding to the right of the range.
    # the rows are expected to remain the same.
    # the first column is expect to remain the same.
    co = ColObj(col)
    rng1 = RangeObj.from_range(rng_name)
    rng2 = rng1 + co
    assert str(rng2) == expected


@pytest.mark.parametrize(
    ("rng_name", "col", "expected"),
    [
        ("B1:C3", "A", "A1:C3"),
        ("D2:E2", "B", "B2:E2"),
        ("R2:V22", "J", "H2:V22"),
    ],
)
def test_add_range_cols_start_col_obj(rng_name: str, col: str, expected: str) -> None:
    from ooodev.utils.data_type.range_obj import RangeObj
    from ooodev.utils.data_type.col_obj import ColObj

    # adding cols to the start is adding cols to the left of the range.
    # The overall cols are increased by adding to the left of the range.
    # the rows are expected to remain the same.
    # the last column is expect to remain the same.
    co = ColObj(col)
    rng1 = RangeObj.from_range(rng_name)
    rng2 = co + rng1
    assert str(rng2) == expected


@pytest.mark.parametrize(
    ("rng_name", "col", "expected"),
    [
        ("A1:C3", "A", "A1:B3"),
        ("A2:D2", "B", "A2:B2"),
        ("C2:R22", "J", "C2:H22"),
    ],
)
def test_subtract_range_cols_end_col_obj(rng_name: str, col: str, expected: str) -> None:
    from ooodev.utils.data_type.range_obj import RangeObj
    from ooodev.utils.data_type.col_obj import ColObj

    # subtracting cols from the end is removing cols from the right of the range.
    # The overall cols are decreased by removeing from the right of the range.
    # the rows are expected to remain the same.
    # the first column is expect to remain the same.
    co = ColObj(col)
    rng1 = RangeObj.from_range(rng_name)
    rng2 = rng1 - co
    assert str(rng2) == expected


@pytest.mark.parametrize(
    ("rng_name", "col", "expected"),
    [
        ("A1:C3", "A", "B1:C3"),
        ("A2:D2", "B", "C2:D2"),
        ("C2:R22", "J", "M2:R22"),
    ],
)
def test_subtract_range_cols_start_col_obj(rng_name: str, col: str, expected: str) -> None:
    from ooodev.utils.data_type.range_obj import RangeObj
    from ooodev.utils.data_type.col_obj import ColObj

    # subtracting cols from the start is removing cols from the left of the range.
    # The overall cols are decreased by removeing from the left of the range.
    # the rows are expected to remain the same.
    # the last column is expect to remain the same.

    co = ColObj(col)
    rng1 = RangeObj.from_range(rng_name)
    rng2 = co - rng1
    assert str(rng2) == expected


# endregion Test RangeObj add/subtract cols using ColObj


# region test RangeObj add/subtract using CellObj
@pytest.mark.parametrize(
    ("rng_name", "cell_name", "expected"),
    [
        ("A1:A1", "A2", "A1:B3"),
        ("C7:D9", "B3", "C7:F12"),
    ],
)
def test_add_range_cols_end_cell_obj(rng_name: str, cell_name: str, expected: str) -> None:
    from ooodev.utils.data_type.range_obj import RangeObj
    from ooodev.utils.data_type.cell_obj import CellObj

    co = CellObj.from_cell(cell_name)
    rng1 = RangeObj.from_range(rng_name)
    rng2 = rng1 + co
    assert str(rng2) == expected


@pytest.mark.parametrize(
    ("rng_name", "cell_name", "expected"),
    [
        ("C3:D4", "A2", "B1:D4"),
        ("C7:D9", "B3", "A4:D9"),
    ],
)
def test_add_range_cols_start_cell_obj(rng_name: str, cell_name: str, expected: str) -> None:
    from ooodev.utils.data_type.range_obj import RangeObj
    from ooodev.utils.data_type.cell_obj import CellObj

    co = CellObj.from_cell(cell_name)
    rng1 = RangeObj.from_range(rng_name)
    rng2 = co + rng1
    assert str(rng2) == expected


@pytest.mark.parametrize(
    ("rng_name", "cell_name", "expected"),
    [
        ("A1:C3", "A2", "A1:B1"),
        ("C7:D9", "C3", "A6:C7"),
    ],
)
def test_subtract_range_cols_end_cell_obj(rng_name: str, cell_name: str, expected: str) -> None:
    from ooodev.utils.data_type.range_obj import RangeObj
    from ooodev.utils.data_type.cell_obj import CellObj

    co = CellObj.from_cell(cell_name)
    rng1 = RangeObj.from_range(rng_name)
    rng2 = rng1 - co
    assert str(rng2) == expected


@pytest.mark.parametrize(
    ("rng_name", "cell_name", "expected"),
    [
        ("R13:Z33", "A2", "S15:Z33"),
        ("E7:R9", "C3", "H9:R10"),
    ],
)
def test_subtract_range_cols_start_cell_obj(rng_name: str, cell_name: str, expected: str) -> None:
    from ooodev.utils.data_type.range_obj import RangeObj
    from ooodev.utils.data_type.cell_obj import CellObj

    co = CellObj.from_cell(cell_name)
    rng1 = RangeObj.from_range(rng_name)
    rng2 = co - rng1
    assert str(rng2) == expected


def test_combine() -> None:
    from ooodev.utils.data_type.range_obj import RangeObj

    # combining range objects should result in a new range object
    # that contains the smallest start column, the smallest start row, the largest end column, the largest end row

    rng1 = RangeObj(col_start="C", col_end="F", row_start=3, row_end=6, sheet_idx=0)
    rng2 = RangeObj(col_start="C", col_end="F", row_start=1, row_end=2, sheet_idx=0)
    rng3 = rng1 / rng2
    assert str(rng3) == "C1:F6"

    rng2 = RangeObj(col_start="A", col_end="F", row_start=1, row_end=2, sheet_idx=0)
    rng3 = rng1 / rng2
    assert str(rng3) == "A1:F6"

    rng2 = RangeObj(col_start="C", col_end="F", row_start=4, row_end=6, sheet_idx=0)
    rng3 = rng1 / rng2
    assert str(rng3) == "C3:F6"

    rng1 = RangeObj(col_start="A", col_end="G", row_start=3, row_end=1, sheet_idx=0)
    rng2 = RangeObj(col_start="C", col_end="F", row_start=1, row_end=2, sheet_idx=0)
    rng3 = rng1 / rng2
    assert str(rng3) == "A1:G3"

    rng1 = RangeObj(col_start="C", col_end="F", row_start=1, row_end=2, sheet_idx=0)
    rng2 = RangeObj(col_start="A", col_end="G", row_start=3, row_end=1, sheet_idx=0)
    rng3 = rng1 / rng2
    assert str(rng3) == "A1:G3"

    rng1 = RangeObj(col_start="C", col_end="F", row_start=1, row_end=2, sheet_idx=0)
    rng2 = RangeObj(col_start="A", col_end="G", row_start=3, row_end=1, sheet_idx=0)
    rng3 = RangeObj(col_start="A", col_end="H", row_start=3, row_end=7, sheet_idx=0)
    rng4 = rng1 / rng2 / rng3
    assert str(rng4) == "A1:H7"

    rng1 = RangeObj(col_start="C", col_end="F", row_start=2, row_end=2, sheet_idx=0)
    rng2 = RangeObj(col_start="A", col_end="G", row_start=3, row_end=1, sheet_idx=0)
    rng3 = RangeObj(col_start="A", col_end="H", row_start=3, row_end=7, sheet_idx=0)
    rng4 = RangeObj(col_start="A", col_end="B", row_start=1, row_end=7, sheet_idx=0)
    rng5 = rng1 / rng2 / rng3 / rng4
    assert str(rng5) == "A1:H7"

    rng1 = RangeObj(col_start="C", col_end="F", row_start=2, row_end=2, sheet_idx=0)
    rng2 = rng1 / "B2:C4"
    assert str(rng2) == "B2:F4"

    rng1 = RangeObj(col_start="C", col_end="F", row_start=2, row_end=2, sheet_idx=0)
    rng2 = rng1 / "B2:C4" / "a1:d7"
    assert str(rng2) == "A1:F7"

    rng1 = RangeObj(col_start="C", col_end="F", row_start=2, row_end=2, sheet_idx=0)
    rng2 = "B2:C4" / rng1
    assert str(rng2) == "B2:F4"

    rng1 = RangeObj(col_start="C", col_end="F", row_start=2, row_end=2, sheet_idx=0)
    rng2 = "B2:C4" / rng1 / "a1:d7"
    assert str(rng2) == "A1:F7"


def test_combine_errors() -> None:
    from ooodev.utils.data_type.range_obj import RangeObj

    rng1 = RangeObj(col_start="C", col_end="F", row_start=2, row_end=2, sheet_idx=0)

    with pytest.raises(ValueError):
        _ = rng1 / "B2"

    with pytest.raises(ValueError):
        _ = rng1 / "B"

    with pytest.raises(ValueError):
        _ = "B2" / rng1

    with pytest.raises(TypeError):
        _ = "B2:C4" / "a1:d7" / rng1


# endregion test RangeObj add/subtract using CellObj
