from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])
from ooodev.utils.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.calc import Calc


def test_scenarios(loader) -> None:
    doc = Calc.create_doc(loader=loader)
    visible = False
    delay = 0  # 1_000
    if visible:
        GUI.set_visible(is_visible=visible, odoc=doc)
    sheet = Calc.get_sheet(doc=doc, index=0)
    try:
        vals = [[11, 12], ["Test13", "Test14"]]
        Calc.insert_scenario(
            sheet=sheet, range_name="B10:C11", vals=vals, name="First Scenario", comment="1st scenario."
        )
        Calc.set_val(value="=B11+C11", sheet=sheet, cell_name="D10")
        Lo.delay(delay)

        vals[0][0] = "Test21"
        vals[0][1] = "Test22"
        vals[1][0] = 23
        vals[1][1] = 24
        Calc.insert_scenario(
            sheet=sheet, range_name="B10:C11", vals=vals, name="Second Scenario", comment="Visible scenario."
        )
        Lo.delay(delay)

        vals[0][0] = 31
        vals[0][1] = 32
        vals[1][0] = "Test33"
        vals[1][1] = "Test34"

        Calc.insert_scenario(
            sheet=sheet, range_name="B10:C11", vals=vals, name="Third Scenario", comment="Last scenario."
        )

        Lo.delay(delay)
        Calc.apply_scenario(sheet=sheet, name="Second Scenario")
        arr = Calc.get_array(sheet=sheet, range_name="D10:D10")
        assert arr[0][0] == 47.0
        Lo.delay(delay)
    finally:
        Lo.close(closeable=doc, deliver_ownership=False)
