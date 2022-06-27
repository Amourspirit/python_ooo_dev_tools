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


def test_make_column_name():
    name = TableHelper.make_column_name(2)
    assert name == "B"

    name = TableHelper.make_column_name(27)
    assert name == "AA"

    name = TableHelper.make_column_name(100)
    assert name == "CV"

    with pytest.raises(ValueError) as ctx:
        name = TableHelper.make_column_name(0)

    with pytest.raises(ValueError) as ctx:
        name = TableHelper.make_column_name(-11)


def test_make_cell_name():
    name = TableHelper.make_cell_name(row=10, col=3)
    assert name == "C10"

    name = TableHelper.make_cell_name(row=100, col=27)
    assert name == "AA100"

    name = TableHelper.make_cell_name(row=101, col=100)
    assert name == "CV101"

    with pytest.raises(ValueError) as ctx:
        name = TableHelper.make_cell_name(row=0, col=3)

    with pytest.raises(ValueError) as ctx:
        name = TableHelper.make_cell_name(row=2, col=0)


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