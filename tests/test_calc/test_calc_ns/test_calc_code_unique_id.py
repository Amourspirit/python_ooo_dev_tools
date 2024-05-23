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
        id1 = sheet.unique_id
        assert len(id1) > 0
        assert sheet.unique_id == id1

        sheet2 = doc.sheets.insert_sheet("MySheet2")
        id2 = sheet2.unique_id
        assert len(id2) > 0
        assert sheet2.unique_id == id2

        sheet3 = doc.sheets.insert_sheet("My Third Sheet")
        id3 = sheet3.unique_id
        assert len(id3) > 0
        assert sheet3.unique_id == id3

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
        assert sheet.unique_id == id1
        sheet2 = doc.sheets[1]
        assert sheet2.unique_id == id2
        sheet3 = doc.sheets[2]
        assert sheet3.unique_id == id3

        c_name = doc.get_sheet_name_from_unique_id(id2)
        assert c_name == "MySheet2"
    finally:
        if doc is not None:
            doc.close()
