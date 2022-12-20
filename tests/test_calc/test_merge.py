from __future__ import annotations
from typing import List
import pytest

if __name__ == "__main__":
    pytest.main([__file__])
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.calc import Calc


def test_merge(loader, test_headless) -> None:
    doc = Calc.create_doc()
    visible = not test_headless
    delay = 0  # 2000 if visible else 0
    if visible:
        GUI.set_visible(is_visible=visible, doc=doc)
    sheet = Calc.get_sheet(doc=doc)

    try:
        rng1 = Calc.get_range_obj("A1:B5")
        assert Calc.is_merged_cells(sheet, rng1) == False
        Calc.merge_cells(sheet, rng1)
        assert Calc.is_merged_cells(sheet, rng1)

        rng2 = Calc.get_range_obj("C3:G10")
        Calc.merge_cells(sheet, rng2, True)
        assert Calc.is_merged_cells(sheet, rng2)

        Calc.unmerge_cells(sheet, rng1)
        assert Calc.is_merged_cells(sheet, rng1) == False

        Calc.merge_cells(sheet=sheet, range_obj=rng1, center=True)
        assert Calc.is_merged_cells(sheet=sheet, range_obj=rng1)

        Calc.set_val(value="Merged", sheet=sheet, cell_obj=rng1.cell_start)
        Calc.set_val(value="Hello World!", sheet=sheet, cell_obj=rng2.cell_start)

        Calc.unmerge_cells(sheet=sheet, range_obj=rng2)
        assert Calc.is_merged_cells(sheet=sheet, range_obj=rng2) == False

        Lo.delay(delay)
    finally:
        Lo.close(closeable=doc, deliver_ownership=False)
