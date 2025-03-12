from __future__ import annotations
import pytest
from pathlib import Path

if __name__ == "__main__":
    pytest.main([__file__])


def test_unique_id(loader, tmp_path) -> None:
    # Test adding controls to a cell and a range
    # The controls are found when calling "cell.control.current_control" by shape position.
    from ooodev.calc import CalcDoc

    doc = None
    pth = Path(tmp_path, "test_unique_id.ods")
    id1 = ""
    id2 = ""
    id3 = ""
    try:
        doc = CalcDoc.create_doc(loader)
        sheet = doc.sheets[0]
        cell1 = sheet["A1"]
        cell1.value = "sheet1.cellA1"
        id1 = cell1.unique_id
        assert len(id1) > 0
        assert cell1.unique_id == id1

        cell2 = sheet["A2"]
        cell2.value = "sheet1.cellA2"
        id2 = cell2.unique_id
        assert len(id2) > 0
        assert cell2.unique_id == id2
        assert cell1.unique_id != cell2.unique_id

        sheet2 = doc.sheets.insert_sheet("MySheet2")
        cell3 = sheet2["A1"]
        cell3.value = "sheet2.cellA1"
        id3 = cell3.unique_id
        assert len(id3) > 0
        assert cell3.unique_id == id3
        assert cell1.unique_id != cell3.unique_id

        doc.save_doc(pth)

    finally:
        if doc is not None:
            doc.close()

    assert id1
    assert id2
    assert id3
    assert pth.exists()
    assert pth.is_file()
    doc = None
    try:
        doc = CalcDoc.open_doc(pth)
        sheet = doc.sheets[0]
        cell1 = sheet["A1"]
        assert cell1.unique_id == id1
        cell2 = sheet["A2"]
        assert cell2.unique_id == id2

        sheet2 = doc.sheets[1]
        cell3 = sheet2["A1"]
        assert cell3.unique_id == id3

        # delete column a1 from sheet 1
        sheet.delete_row(0)
        cell1 = sheet["A1"]
        assert cell1.value == "sheet1.cellA2"
        assert cell1.unique_id == id2

        sheet2.insert_row(0)
        cell3 = sheet2["A2"]
        assert cell3.value == "sheet2.cellA1"
        assert cell3.unique_id == id3
    finally:
        if doc is not None:
            doc.close()
