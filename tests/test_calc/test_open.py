from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.calc import Calc


def test_oopen_small_totals(copy_fix_calc, loader) -> None:
    doc_path = copy_fix_calc("small_totals.ods")
    doc = Calc.open_doc(fnm=str(doc_path), loader=loader)
    assert doc is not None, "Could not open small_totals.ods"
    visible = False
    if visible:
        GUI.set_visible(is_visible=visible, odoc=doc)
    sheet = Calc.get_sheet(doc=doc)
    val = Calc.get_val(sheet=sheet, cell_name="A1")
    assert val == "Stud. No."

    Lo.close(closeable=doc, deliver_ownership=False)


def test_open_small_totals_no_loader(copy_fix_calc, loader) -> None:
    doc_path = copy_fix_calc("small_totals.ods")
    doc = Calc.open_doc(fnm=str(doc_path))
    assert doc is not None, "Could not open small_totals.ods"
    visible = False
    if visible:
        GUI.set_visible(is_visible=visible, odoc=doc)
    sheet = Calc.get_sheet(doc=doc)
    val = Calc.get_val(sheet=sheet, cell_name="A1")
    assert val == "Stud. No."

    Lo.close(closeable=doc, deliver_ownership=False)


def test_open_no_file_no_loader(loader) -> None:
    doc = Calc.open_doc()
    assert doc is not None, "Could not open new calc document"
    visible = False
    if visible:
        GUI.set_visible(is_visible=visible, odoc=doc)
    sheet = Calc.get_sheet(doc=doc)
    Calc.set_val(value="Stud. No.", sheet=sheet, cell_name="A1")
    val = Calc.get_val(sheet=sheet, cell_name="A1")
    assert val == "Stud. No."
    Lo.close(closeable=doc, deliver_ownership=False)
