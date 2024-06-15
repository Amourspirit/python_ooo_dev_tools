from __future__ import annotations
import pytest

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.calc import CalcDoc
from ooodev.calc import CalcCell


def test_cell_from_obj(loader):
    doc = None
    try:
        doc = CalcDoc.create_doc(loader=loader)
        sheet = doc.sheets[0]

        cell = sheet["A1"]
        cp = cell.component
        co = CalcCell.from_obj(cp)
        assert co is not None
        assert isinstance(co, CalcCell)
        assert co.cell_obj == cell.cell_obj
        assert co.calc_doc.runtime_uid == doc.runtime_uid
        assert sheet.unique_id == co.calc_sheet.unique_id

    finally:
        if doc is not None:
            doc.close()
