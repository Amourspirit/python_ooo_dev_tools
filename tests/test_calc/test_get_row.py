from ooodev.loader.lo import Lo
from ooodev.office.calc import Calc


def test_get_row(loader) -> None:

    doc = Calc.create_doc(loader)
    assert doc is not None
    try:
        sheet = Calc.get_sheet(doc=doc, index=0)

        # test Empty Sheet
        new_vals = Calc.get_row(sheet=sheet, row_idx=0)
        assert len(new_vals) == 1

        new_vals = Calc.get_row(sheet, 0)
        assert len(new_vals) == 1

        new_vals = Calc.get_row(sheet=sheet, row_idx=10)
        assert len(new_vals) == 0

        # get the whole used area of sheet
        cell_range = Calc.find_used_range(sheet)
        # gets the first column
        new_vals = Calc.get_row(cell_range)
        assert len(new_vals) == 1  # A1:A1

        # test with data
        vals_len = 12

        vals = [float(i) for i in range(vals_len)]
        # set_col(sheet: XSpreadsheet, values: Sequence[Any], cell_name: str)
        # keyword arguments
        Calc.set_row(sheet=sheet, values=vals, cell_name="A2")
        range_name = f"{Calc.get_cell_str(col=0, row=1)}:{Calc.get_cell_str(col=vals_len -1, row=1)}"
        new_vals = Calc.get_row(sheet=sheet, range_name=range_name)
        for i, val in enumerate(vals):
            assert new_vals[i] == val

        # positional
        new_vals = Calc.get_row(sheet, range_name)
        assert len(new_vals) == vals_len

        new_vals = Calc.get_row(sheet=sheet, row_idx=0)
        assert len(new_vals) == 0

        new_vals = Calc.get_row(sheet=sheet, row_idx=1)
        assert len(new_vals) == vals_len

        # get the whole used area of sheet
        cell_range = Calc.find_used_range(sheet)
        # get the first row of the range
        # in this case the row at index 0 is not used
        # can expect cellrange to be starting at row index 1
        new_vals = Calc.get_row(cell_range)
        assert len(new_vals) == vals_len

        vals = [float(i * i) for i in range(vals_len)]

        # add data to the first row
        Calc.set_row(sheet=sheet, values=vals, cell_name="A1")
        # get the whole used area of sheet
        cell_range = Calc.find_used_range(sheet)
        # get the first row
        new_vals = Calc.get_row(cell_range)
        assert len(new_vals) == vals_len
        assert new_vals[10] == 100.0

    finally:
        Lo.close(closeable=doc, deliver_ownership=False)
