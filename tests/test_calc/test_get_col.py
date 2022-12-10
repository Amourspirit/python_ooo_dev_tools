from ooodev.utils.lo import Lo
from ooodev.office.calc import Calc


def test_get_col(loader) -> None:

    doc = Calc.create_doc(loader)
    assert doc is not None
    try:
        sheet = Calc.get_sheet(doc=doc, index=0)

        # test Empty Sheet
        new_vals = Calc.get_col(sheet=sheet, col_name="a")
        assert len(new_vals) == 1

        new_vals = Calc.get_col(sheet, 0)
        assert len(new_vals) == 1

        new_vals = Calc.get_col(sheet=sheet, col_idx=10)
        assert len(new_vals) == 0

        # get the whole used area of sheet
        cell_range = Calc.find_used_range(sheet)
        # gets the first column
        new_vals = Calc.get_col(cell_range)
        assert len(new_vals) == 1  # A1:A1

        # test with data
        vals_len = 12

        vals = [float(i) for i in range(vals_len)]
        # set_col(sheet: XSpreadsheet, values: Sequence[Any], cell_name: str)
        # keyword arguments
        Calc.set_col(sheet=sheet, values=vals, cell_name="A1")
        range_name = f"{Calc.get_cell_str(col=0, row=0)}:{Calc.get_cell_str(col=0, row= vals_len -1)}"
        new_vals = Calc.get_col(sheet=sheet, range_name=range_name)
        for i, val in enumerate(vals):
            assert new_vals[i] == val

        # positional
        new_vals = Calc.get_col(sheet, range_name)
        assert len(new_vals) == vals_len

        new_vals = Calc.get_col(sheet=sheet, col_name="A")
        assert len(new_vals) == vals_len

        # positional
        new_vals = Calc.get_col(sheet, "A")
        assert len(new_vals) == vals_len

        new_vals = Calc.get_col(sheet=sheet, col_idx=0)
        assert len(new_vals) == vals_len

        # positional
        new_vals = Calc.get_col(sheet, 0)
        assert len(new_vals) == vals_len

        new_vals = Calc.get_col(sheet=sheet, col_name="b")
        assert len(new_vals) == 0

        # get the whole used area of sheet
        cell_range = Calc.find_used_range(sheet)
        # get the first column
        new_vals = Calc.get_col(cell_range)
        assert len(new_vals) == vals_len

    finally:
        Lo.close(closeable=doc, deliver_ownership=False)
