from __future__ import annotations
import pytest

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.calc import CalcDoc
from ooodev.calc import CalcCellRange


def test_cell_range_from_obj(loader):
    doc = None
    try:
        doc = CalcDoc.create_doc(loader=loader)
        sheet = doc.sheets[0]

        rng = doc.range_converter.get_range_obj("A1:B3")

        rng_obj = sheet.get_range(range_obj=rng)

        cp = rng_obj.component
        co = CalcCellRange.from_obj(cp)
        assert co is not None
        assert isinstance(co, CalcCellRange)
        assert co.range_obj == rng_obj.range_obj
        assert co.calc_doc.runtime_uid == doc.runtime_uid
        assert sheet.unique_id == co.calc_sheet.unique_id

    finally:
        if doc is not None:
            doc.close()
