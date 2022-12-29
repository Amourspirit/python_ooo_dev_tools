import pytest
from typing import cast

if __name__ == "__main__":
    pytest.main([__file__])


def test_cell_obj(loader) -> None:
    from ooodev.utils.data_type.cell_obj import CellObj
    from ooodev.utils.data_type.range_obj import RangeObj

    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    doc = Calc.create_doc()
    sheet = Calc.get_sheet(doc)

    try:

        name = "B22"
        addr = Calc.get_cell_address(sheet, name)
        cell = CellObj.from_cell(addr)
        assert cell.col == "B"
        assert cell.row == 22

        name = "r75"
        cell = CellObj.from_cell(name)
        assert cell.col == "R"
        assert cell.row == 75
        assert cell.col_info.value == "R"
        assert cell.col_info.index == 17
        assert cell.row_info.value == 75
        assert cell.row_info.index == 74
        assert cell == "R75"
        assert cell == "r75"

        rng = RangeObj.from_range("a1:r75")
        assert rng.cell_end == cell

        addr = cell.get_cell_address()
        assert addr.Column == 17
        assert addr.Row == 74
        assert addr.Sheet == 0

        cell = CellObj.from_idx(1, 7)
        assert cell.sheet_idx == 0
        assert cell.col == "B"
        assert cell.row == 8

    finally:
        Lo.close_doc(doc)


def test_cell_obj_init_error(loader) -> None:
    from ooodev.utils.data_type.cell_obj import CellObj

    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    doc = Calc.create_doc()
    _ = Calc.get_sheet(doc)

    try:
        with pytest.raises(AssertionError):
            _ = CellObj("@", 2)

        with pytest.raises(AssertionError):
            _ = CellObj("B:", 2)

        with pytest.raises(AssertionError):
            _ = CellObj("A", 0)

        with pytest.raises(AssertionError):
            _ = CellObj("A", -1)

        with pytest.raises(AssertionError):
            _ = CellObj("", 10)

        with pytest.raises(AssertionError):
            _ = CellObj.from_cell("B!2")

        with pytest.raises(AssertionError):
            _ = CellObj.from_cell("2")

        with pytest.raises(ValueError):
            _ = CellObj.from_cell("B:a2")

        with pytest.raises(ValueError):
            _ = CellObj.from_cell("b.")

        with pytest.raises(ValueError):
            _ = CellObj.from_cell("b:0")

        with pytest.raises(ValueError):
            _ = CellObj.from_idx(-1, 0, 0)

        with pytest.raises(AssertionError):
            _ = CellObj.from_idx(0, -1, 0)

    finally:
        Lo.close_doc(doc)


def test_cell_obj_to_cell_values(loader) -> None:
    from ooodev.utils.data_type.cell_obj import CellObj

    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    doc = Calc.create_doc()
    _ = Calc.get_sheet(doc)

    try:

        name = "B2"
        cell1 = CellObj.from_cell(name)
        cv1 = cell1.get_cell_values()
        assert cv1.col == 1
        assert cv1.row == 1
        assert cv1.sheet_idx == cell1.sheet_idx

        cell2 = CellObj.from_cell(cv1)
        assert cell2 == cell1
        cv2 = cell2.get_cell_values()
        assert cv2 == cv1
        assert cv2 == name

    finally:
        Lo.close_doc(doc)


def test_cell_adjacent(loader) -> None:
    from ooodev.utils.data_type.cell_obj import CellObj

    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    doc = Calc.create_doc()
    _ = Calc.get_sheet(doc)

    try:
        c3 = CellObj.from_cell("C3")
        b3 = c3.left
        assert b3.col == "B"
        assert b3.row == c3.row

        d3 = c3.right
        assert d3.col == "D"
        assert d3.row == c3.row

        e4 = d3.down.right
        assert e4.col == "E"
        assert e4.row == 4

        a3 = c3.left.left
        assert a3.col == "A"
        assert a3.row == c3.row

        a1 = c3.left.left.up.up
        assert a1.col == "A"
        assert a1.row == 1
    finally:
        Lo.close_doc(doc)


def test_cell_left_up_error(loader) -> None:
    from ooodev.utils.data_type.cell_obj import CellObj

    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    doc = Calc.create_doc()
    _ = Calc.get_sheet(doc)

    try:
        b2 = CellObj.from_cell("B2")
        with pytest.raises(IndexError):
            _ = b2.left.left

        with pytest.raises(IndexError):
            _ = b2.left.left.left

        with pytest.raises(IndexError):
            _ = b2.up.up

        with pytest.raises(IndexError):
            _ = b2.up.up.up

    finally:
        Lo.close_doc(doc)


def test_cell_math(loader) -> None:
    from ooodev.utils.data_type.cell_obj import CellObj

    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    doc = Calc.create_doc()
    _ = Calc.get_sheet(doc)

    try:
        b2 = CellObj.from_cell("B2")

        b4 = b2 + 2
        assert b4.col == "B"
        assert b4.row == 4

        b4 = cast(CellObj, b2 + (b2.row_info + 2))
        assert b4.col == "B"
        assert b4.row == 4

        d2 = cast(CellObj, b2 + (b2.col_info + 2))
        assert d2.col == "D"
        assert d2.row == 2

        e2 = cast(CellObj, b2 + "C")  # add 3 columns
        assert e2.col == "E"
        assert e2.row == 2

        a2 = cast(CellObj, e2 - "D")  # subtract 4 columns
        assert a2.col == "A"
        assert a2.row == 2

        a2 = cast(CellObj, e2 - (e2.col_info - 4))
        assert a2.col == "A"
        assert a2.row == 2

        b2 = b4 - 2
        assert b2.col == "B"
        assert b2.row == 2

        e5 = cast(CellObj, e2 + (e2.row_info + 3))
        assert e5.col == "E"
        assert e5.row == 5

        e3 = cast(CellObj, e5 - (e5.row_info - 2))
        assert e3.col == "E"
        assert e3.row == 3

        # add two cells and two colums by adding CellObj instnaces
        g5 = cast(CellObj, e3 + CellObj("B", 2, sheet_idx=e3.sheet_idx))
        assert g5.col == "G"
        assert g5.row == 5

        # subtract to cells and tow columns by subtracting CellObj instance
        e3 = cast(CellObj, g5 - CellObj("B", 2, sheet_idx=e3.sheet_idx))
        assert e3.col == "E"
        assert e3.row == 3

    finally:
        Lo.close_doc(doc)


def test_cell_math_errors(loader) -> None:
    from ooodev.utils.data_type.cell_obj import CellObj

    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    doc = Calc.create_doc()
    _ = Calc.get_sheet(doc)
    try:
        b2 = CellObj.from_cell("B2")
        with pytest.raises(IndexError):
            _ = b2 - 2

        with pytest.raises(IndexError):
            _ = b2 - (b2.row_info - 2)

        with pytest.raises(IndexError):
            _ = b2 - (b2.row_info - 2)

        with pytest.raises(IndexError):
            _ = b2 - (b2.col_info - 2)
    finally:
        Lo.close_doc(doc)
