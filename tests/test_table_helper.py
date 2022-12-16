from typing import Any
import pytest

if __name__ == "__main__":
    pytest.main([__file__])
from ooodev.utils.table_helper import TableHelper


def test_create_2d():
    num_rows = 5
    num_cols = 10
    arr = TableHelper.make_2d_array(num_rows=num_rows, num_cols=num_cols)
    assert len(arr) == num_rows, "Number of rows does not match"
    assert len(arr[0]) == num_cols, "Number of cols does not match"

    num_rows = 5
    num_cols = 10
    arr = TableHelper.make_2d_list(num_rows=num_rows, num_cols=num_cols)
    assert len(arr) == num_rows, "Number of rows does not match"
    assert len(arr[0]) == num_cols, "Number of cols does not match"

    num_rows = 22
    num_cols = 8
    arr = TableHelper.make_2d_array(num_rows=num_rows, num_cols=num_cols, val="a")
    assert len(arr) == num_rows, "Number of rows does not match"
    assert len(arr[0]) == num_cols, "Number of cols does not match"

    for row in arr:
        for col in row:
            assert col == "a"


@pytest.mark.parametrize(
    ("row", "col", "expected"), [(1, 1, "A1"), (10, 26, "Z10"), (10, 3, "C10"), (101, 100, "CV101")]
)
def test_make_cell_name(row: int, col: int, expected: str) -> None:
    name = TableHelper.make_cell_name(row=row, col=col)
    assert name == expected


@pytest.mark.parametrize(
    ("row", "col", "expected"), [(0, 0, "A1"), (1, 1, "B2"), (9, 25, "Z10"), (9, 2, "C10"), (100, 99, "CV101")]
)
def test_make_cell_name_zero(row: int, col: int, expected: str) -> None:
    name = TableHelper.make_cell_name(row=row, col=col, zero_index=True)
    assert name == expected


def test_make_cell_name_error() -> None:
    with pytest.raises(ValueError):
        _ = TableHelper.make_cell_name(row=0, col=3)

    with pytest.raises(ValueError):
        _ = TableHelper.make_cell_name(row=2, col=0)

    with pytest.raises(ValueError):
        _ = TableHelper.make_cell_name(row=2, col=-1, zero_index=True)

    with pytest.raises(ValueError):
        _ = TableHelper.make_cell_name(row=-2, col=0, zero_index=True)


@pytest.mark.parametrize(("val", "expected"), [("A", 1), ("C2", 3), ("AA2", 27)])
def test_col_name_to_int(val: int, expected: str) -> None:
    i = TableHelper.col_name_to_int(val)
    assert i == expected


@pytest.mark.parametrize(("val", "expected"), [("A", 0), ("C2", 2), ("AA2", 26)])
def test_col_name_to_int_zero_index(val: int, expected: str) -> None:
    i = TableHelper.col_name_to_int(val, True)
    assert i == expected


@pytest.mark.parametrize(("val", "expected"), [("1", 1), ("C2", 2), ("AA27", 27), ("abz27756", 27756)])
def test_row_name_to_int(val: int, expected: str) -> None:
    i = TableHelper.row_name_to_int(val)
    assert i == expected


@pytest.mark.parametrize(("val", "expected"), [("1", 0), ("C2", 1), ("AA27", 26), ("abz27756", 27755)])
def test_row_name_to_int_zero(val: int, expected: str) -> None:
    i = TableHelper.row_name_to_int(val, True)
    assert i == expected


@pytest.mark.parametrize(("val"), [("AA-27"), ("-1")])
def test_row_name_to_int_err(val: int) -> None:
    with pytest.raises(ValueError):
        _ = TableHelper.row_name_to_int(val)


@pytest.mark.parametrize(("val", "expected"), [(1, "A"), (2, "B"), (27, "AA"), (100, "CV")])
def test_make_column_name(val: int, expected: str) -> None:
    name = TableHelper.make_column_name(val)
    assert name == expected


@pytest.mark.parametrize(("val", "expected"), [(0, "A"), (1, "B"), (26, "AA"), (99, "CV")])
def test_make_column_name_zero_based(val: int, expected: str) -> None:
    name = TableHelper.make_column_name(val, True)
    assert name == expected


def test_make_column_name_error() -> None:

    with pytest.raises(ValueError):
        _ = TableHelper.make_column_name(0)

    with pytest.raises(ValueError):
        _ = TableHelper.make_column_name(-1, True)

    with pytest.raises(ValueError):
        _ = TableHelper.make_column_name(-11)


def test_to_list():
    obj = ("one",)
    lst = TableHelper.to_list(obj)
    assert isinstance(lst, list)
    for i in range(len(obj)):
        assert lst[i] == obj[i]

    obj = ("one", "two", "three", 4, 5, "six")
    lst = TableHelper.to_list(obj)
    for i in range(len(obj)):
        assert lst[i] == obj[i]

    obj_lst = list(obj)
    lst = TableHelper.to_list(obj_lst)
    for i in range(len(obj_lst)):
        assert lst[i] == obj_lst[i]


def test_to_tuple():
    obj = ["one"]
    tpl = TableHelper.to_tuple(obj)
    assert isinstance(tpl, tuple)
    for i in range(len(tpl)):
        assert obj[i] == tpl[i]

    obj = ["one", "two", "three", 4, 5, "six"]
    tpl = TableHelper.to_tuple(obj)
    for i in range(len(tpl)):
        assert tpl[i] == obj[i]

    obj_lst = tuple(obj)
    tpl = TableHelper.to_tuple(obj_lst)
    for i in range(len(obj_lst)):
        assert tpl[i] == obj_lst[i]


def test_to_2d_list():
    obj = (("one",),)
    lst = TableHelper.to_2d_list(obj)
    assert isinstance(lst, list)
    assert len(lst) == 1
    assert isinstance(lst[0], list)
    assert lst[0][0] == "one"

    obj = (("one", "two", "three"), ("a", "b", "c"), (1, 2, 3))
    lst = TableHelper.to_2d_list(obj)
    assert isinstance(lst, list)
    assert len(lst) == 3
    for row in lst:
        assert isinstance(row, list)
    assert lst[0][2] == "three"
    assert lst[1][1] == "b"
    assert lst[2][0] == 1

    obj = [("one", "two", "three"), ["a", "b", "c"], (1, 2, 3)]
    lst = TableHelper.to_2d_list(obj)
    assert isinstance(lst, list)
    assert len(lst) == 3
    for row in lst:
        assert isinstance(row, list)
    assert lst[0][2] == "three"
    assert lst[1][1] == "b"
    assert lst[2][0] == 1

    #  1-row 2D array
    obj = ("one", "two", "three")
    tpl = TableHelper.to_2d_list(obj)
    assert isinstance(tpl, list)
    assert len(tpl) == 1
    assert isinstance(tpl[0], list)
    assert tpl[0][0] == "one"


def test_to_2d_tuple():
    obj = [["one"]]
    tpl = TableHelper.to_2d_tuple(obj)
    assert isinstance(tpl, tuple)
    assert len(tpl) == 1
    assert isinstance(tpl[0], tuple)
    assert tpl[0][0] == "one"

    obj = [["one", "two", "three"], ["a", "b", "c"], [1, 2, 3]]
    tpl = TableHelper.to_2d_tuple(obj)
    assert isinstance(tpl, tuple)
    assert len(tpl) == 3
    for row in tpl:
        assert isinstance(row, tuple)
    assert tpl[0][2] == "three"
    assert tpl[1][1] == "b"
    assert tpl[2][0] == 1

    obj = [("one", "two", "three"), ["a", "b", "c"], (1, 2, 3)]
    tpl = TableHelper.to_2d_tuple(obj)
    assert isinstance(tpl, tuple)
    assert len(tpl) == 3
    for row in tpl:
        assert isinstance(row, tuple)
    assert tpl[0][2] == "three"
    assert tpl[1][1] == "b"
    assert tpl[2][0] == 1

    #  1-row 2D array
    obj = ["one", "two", "three"]
    tpl = TableHelper.to_2d_tuple(obj)
    assert isinstance(tpl, tuple)
    assert len(tpl) == 1
    assert isinstance(tpl[0], tuple)
    assert tpl[0][0] == "one"


def test_col_name_to_int():
    name = "a"
    i = TableHelper.col_name_to_int(name)
    assert i == 1

    name = "A"
    i = TableHelper.col_name_to_int(name)
    assert i == 1

    name = "Z"
    i = TableHelper.col_name_to_int(name)
    assert i == 26

    name = "AA"
    i = TableHelper.col_name_to_int(name)
    assert i == 27

    name = "ABC"
    i = TableHelper.col_name_to_int(name)
    assert i == 731

    col_name = TableHelper.make_column_name(473)
    i = TableHelper.col_name_to_int(col_name)
    assert i == 473


def test_table_2d_to_dict(bond_movies_table) -> None:
    dt = TableHelper.table_2d_to_dict(bond_movies_table)
    assert len(dt) == 24
    row = dt[0]
    assert row["Title"] == "Dr. No"
    assert row["Year"] == "1962"
    assert row["Actor"] == "Sean Connery"
    assert row["Director"] == "Terence Young"

    row = dt[11]

    assert row["Title"] == "For Your Eyes Only"
    assert row["Year"] == "1981"
    assert row["Actor"] == "Roger Moore"
    assert row["Director"] == "John Glen"

    row = dt[23]

    assert row["Title"] == "Spectre"
    assert row["Year"] == "2015"
    assert row["Actor"] == "Daniel Craig"
    assert row["Director"] == "Sam Mendes"

    with pytest.raises(ValueError):
        TableHelper.table_2d_to_dict([])


def test_table_dict_to_table(bond_movies_lst_dict) -> None:
    tbl = TableHelper.table_dict_to_table(bond_movies_lst_dict)
    assert len(tbl) == 25
    cols = tbl[0]
    assert cols[0] == "Title"
    assert cols[1] == "Year"
    assert cols[2] == "Actor"
    assert cols[3] == "Director"

    tbl[12][0] == "For Your Eyes Only"
    tbl[12][1] == "1981"
    tbl[12][2] == "Roger Moore"
    tbl[12][3] == "John Glen"

    tbl[24][0] == "Spectre"
    tbl[24][1] == "2015"
    tbl[24][2] == "Daniel Craig"
    tbl[24][3] == "Sam Mendes"

    with pytest.raises(ValueError):
        TableHelper.table_dict_to_table([])


@pytest.mark.parametrize(
    "input,expected",
    [
        ((-2, 4, 7), -2),
        ((200, "", 56, "", 17), 17),
        ([600_000, 1_000_000, 200_001], 200_001),
        (
            (
                (500, 700, 2, 19_000, 7455),
                (-500, -700, 1, -19_000, -7455),
            ),
            -19_000,
        ),
    ],
)
def test_get_sm_int(input, expected) -> None:
    result = TableHelper.get_smallest_int(input)
    assert result == expected


@pytest.mark.parametrize(
    "input,expected",
    [
        ((-2, 4, 7), 7),
        ((200, "", 56, "", 17), 200),
        ([600_000, 1_000_000, 200_001], 1_000_000),
        (
            (
                (500, 700, 2, 19_000, 7455),
                (-500, -700, 1, -19_000, -7455),
            ),
            19_000,
        ),
    ],
)
def test_get_largest_int(input, expected) -> None:
    result = TableHelper.get_largest_int(input)
    assert result == expected


@pytest.mark.parametrize(
    "input,expected",
    [
        (("this", "that", "is"), 2),
        (["", "fasting is slimming", "I am not sure about that"], 0),
        ((56,), -1),
        (
            (
                ("once", "upon", "a", "time"),
                ("there", "was", "a", "big"),
                ("bad", "wolf", "in", "the enchantede forest"),
            ),
            1,
        ),
    ],
)
def test_get_sm_str(input, expected) -> None:
    result = TableHelper.get_smallest_str(input)
    assert result == expected


@pytest.mark.parametrize(
    "input,expected",
    [
        (("this", "that", "is"), 4),
        (["", "fasting is slimming", "I am not sure about that"], 24),
        ((56,), -1),
        (
            (
                ("once", "upon", "a", "time"),
                ("there", "was", "a", "big"),
                ("bad", "wolf", "in", "the enchantede forest"),
            ),
            21,
        ),
    ],
)
def test_get_large_str(input, expected) -> None:
    result = TableHelper.get_largest_str(input)
    assert result == expected


def test_get_largest_int_empty_errs() -> None:
    with pytest.raises(ValueError):
        TableHelper.get_largest_int([])


def test_get_smallest_int_empty_errs() -> None:
    with pytest.raises(ValueError):
        TableHelper.get_smallest_int([])


@pytest.mark.parametrize(
    "lst_len,col_count,empty_val,expected_len", [(3, 3, "", 1), (2, 3, None, 1), (1, 3, None, 1), (0, 3, -1, 1)]
)
def test_convert_1d_to_2d(lst_len: int, col_count: int, empty_val: Any, expected_len: int) -> None:
    lst = list(range(lst_len))
    result = TableHelper.convert_1d_to_2d(seq_obj=lst, col_count=col_count, empty_cell_val=empty_val)
    assert len(result) == expected_len
    for row in result:
        assert len(row) == col_count


@pytest.mark.parametrize(
    "lst_len,col_count,expected_len,expected_last_row_len",
    [(5, 4, 2, 1), (38, 7, 6, 3), (50, 2, 25, 2), (50, 1, 50, 1)],
)
def test_convert_1d_to_2d_no_fill(lst_len: int, col_count: int, expected_len: int, expected_last_row_len: int) -> None:
    lst = list(range(lst_len))
    result = TableHelper.convert_1d_to_2d(seq_obj=lst, col_count=col_count)
    assert len(result) == expected_len
    lst_row = result[-1:][0]
    assert len(lst_row) == expected_last_row_len


def test_convert_1d_to_2d_col_error() -> None:
    with pytest.raises(ValueError):
        TableHelper.convert_1d_to_2d([1, 3], 0)


def test_get_range_values() -> None:
    from ooodev.utils.data_type.range_values import RangeValues

    rng_name = "A2:D6"
    rv1 = TableHelper.get_range_values(rng_name)
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

    assert ro.start.col_info.index == 0
    assert ro.start.col_info.value == "A"
    assert ro.start.row_info.index == 1
    assert ro.start.row_info.value == 2

    assert ro.end.col_info.index == 3
    assert ro.end.col_info.value == "D"
    assert ro.end.row_info.index == 5
    assert ro.end.row_info.value == 6

    assert ro.col_start == "A"
    assert ro.col_end == "D"
    assert ro.row_start == 2
    assert ro.row_end == 6

    assert ro.start.col_info < ro.end.col_info
    assert ro.start.row_info < ro.end.row_info

    assert ro.end.col_info > ro.start.col_info
    assert ro.end.row_info > ro.start.row_info

    assert ro.start.col_info != ro.end.col_info
    assert ro.start.row_info != ro.end.row_info


def test_get_range_obj() -> None:
    from ooodev.utils.data_type.range_obj import RangeObj

    ro1 = TableHelper.get_range_obj(range_name="A2:D6")
    assert ro1.col_start == "A"
    assert ro1.row_start == 2
    assert ro1.col_end == "D"
    assert ro1.row_end == 6

    ro2 = TableHelper.get_range_obj(range_name="a2:d6")
    assert ro2.col_start == "A"
    assert ro2.row_start == 2
    assert ro2.col_end == "D"
    assert ro2.row_end == 6

    assert ro1 == ro2
    assert ro1.start.col_info.index == 0
    assert ro1.end.col_info.index == 3

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
