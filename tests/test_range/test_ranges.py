from typing import Any
import pytest

if __name__ == "__main__":
    pytest.main([__file__])
from ooodev.utils.table_helper import TableHelper


def test_get_range_values(loader) -> None:
    from ooodev.utils.data_type.range_values import RangeValues
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    doc = Calc.create_doc()
    # sheet = Calc.get_sheet(doc)

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
    finally:
        Lo.close_doc(doc)


def test_get_range_obj(loader) -> None:
    from ooodev.utils.data_type.range_obj import RangeObj
    from ooodev.utils.lo import Lo
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

        assert ro1.cell_start.col_info.cell_obj is ro1.cell_start
        assert ro1.cell_start.row_info.cell_obj is ro1.cell_start

        assert ro1.cell_end.col_info.cell_obj is ro1.cell_end
        assert ro1.cell_end.row_info.cell_obj is ro1.cell_end

        assert ro1.cell_start.col == "A"
        assert ro1.cell_start.row == 2

        assert ro1.cell_end.col == "D"
        assert ro1.cell_end.row == 6

        assert ro1.cell_start.col_info.value == "A"
        assert ro1.cell_start.col_info.index == 0

        assert ro1.cell_start.row_info.value == 2
        assert ro1.cell_start.row_info.index == 1

        assert ro1.cell_end.col_info.value == "D"
        assert ro1.cell_end.col_info.index == 3

        assert ro1.cell_end.row_info.value == 6
        assert ro1.cell_end.row_info.index == 5

        assert ro1.cell_start.sheet_idx == ro1.sheet_idx
        assert ro1.cell_end.sheet_idx == ro1.sheet_idx

        ro2 = RangeObj.from_range(range_val="a2:d6")
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

        ro3 = RangeObj.from_range(rv2)
        assert ro3 == ro2
    finally:
        Lo.close_doc(doc)


def test_get_range_values_cell_address(loader) -> None:
    from ooodev.utils.data_type.range_values import RangeValues
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    # from ooo.dyn.table.cell_range_address import CellRangeAddress

    doc = Calc.create_doc()
    # sheet = Calc.get_sheet(doc)

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

    # from ooodev.utils.data_type.range_values import RangeValues
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    # from ooo.dyn.table.cell_range_address import CellRangeAddress

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
