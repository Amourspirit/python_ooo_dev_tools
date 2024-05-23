from __future__ import annotations
import pytest
from pathlib import Path

if __name__ == "__main__":
    pytest.main([__file__])


def test_code_name(loader, tmp_path) -> None:
    # Test adding controls to a cell and a range
    # The controls are found when calling "cell.control.current_control" by shape position.
    from ooodev.calc import CalcDoc

    doc = None
    pth = Path(tmp_path, "test_code_name.ods")
    code_name = ""
    code_name2 = ""
    code_name3 = ""
    try:
        doc = CalcDoc.create_doc(loader)
        sheet = doc.sheets[0]
        code_name = sheet.code_name
        assert len(code_name) > 0
        sheet.name = "MySheet"

        assert sheet.code_name == code_name
        sheet2 = doc.sheets.insert_sheet("MySheet2")
        code_name2 = sheet2.code_name
        sheet2.name = "My Other Sheet"
        assert sheet2.code_name == code_name2

        sheet3 = doc.sheets.insert_sheet("My Third Sheet")
        code_name3 = sheet3.code_name
        sheet3.name = "Third Sheet"
        assert sheet3.code_name == code_name3

        doc.save_doc(pth)

    finally:
        if doc is not None:
            doc.close()

    assert code_name
    assert code_name2
    assert code_name3
    assert pth.exists()
    assert pth.is_file()
    doc = None
    try:
        doc = CalcDoc.open_doc(pth)
        sheet = doc.sheets[0]
        assert sheet.code_name == code_name
        sheet2 = doc.sheets[1]
        assert sheet2.code_name == code_name2
        sheet3 = doc.sheets[2]
        assert sheet3.code_name == code_name3

        c_name = doc.get_sheet_name_from_code_name(code_name)
        assert c_name == "MySheet"

    finally:
        if doc is not None:
            doc.close()
