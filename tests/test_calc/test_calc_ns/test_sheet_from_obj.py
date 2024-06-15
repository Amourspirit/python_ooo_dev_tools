from __future__ import annotations
import pytest

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.calc import CalcDoc
from ooodev.calc import CalcSheet


def test_sheet_from_obj_sheet(loader):
    doc = None
    try:
        doc = CalcDoc.create_doc(loader=loader)
        sheet = doc.sheets[0]

        cp = sheet.component
        co = CalcSheet.from_obj(cp)
        assert co is not None
        assert isinstance(co, CalcSheet)
        assert co.name == sheet.name
        assert co.calc_doc.runtime_uid == doc.runtime_uid

    finally:
        if doc is not None:
            doc.close()


def test_sheet_from_obj_cell(loader):
    doc = None
    try:
        doc = CalcDoc.create_doc(loader=loader)
        doc.runtime_uid
        sheet = doc.sheets[0]
        calc_cell = sheet["A1"]

        co = CalcSheet.from_obj(calc_cell)
        assert co is not None
        assert isinstance(co, CalcSheet)
        assert co.name == sheet.name
        assert co.calc_doc.runtime_uid == doc.runtime_uid

        co = CalcSheet.from_obj(calc_cell.component)
        assert co is not None
        assert isinstance(co, CalcSheet)
        assert co.name == sheet.name
        assert co.calc_doc.runtime_uid == doc.runtime_uid

    finally:
        if doc is not None:
            doc.close()


def test_sheet_from_obj_range(loader):
    doc = None
    try:
        doc = CalcDoc.create_doc(loader=loader)
        sheet = doc.sheets[0]
        rng = doc.range_converter.get_range_obj("A1:B3")

        rng_obj = sheet.get_range(range_obj=rng)

        co = CalcSheet.from_obj(rng_obj)
        assert co is not None
        assert isinstance(co, CalcSheet)
        assert co.name == sheet.name
        assert co.calc_doc.runtime_uid == doc.runtime_uid

        co = CalcSheet.from_obj(rng_obj.component)
        assert co is not None
        assert isinstance(co, CalcSheet)
        assert co.name == sheet.name
        assert co.calc_doc.runtime_uid == doc.runtime_uid

    finally:
        if doc is not None:
            doc.close()
