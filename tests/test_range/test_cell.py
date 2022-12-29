import pytest

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
