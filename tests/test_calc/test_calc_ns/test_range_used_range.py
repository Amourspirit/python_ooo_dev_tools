from __future__ import annotations
import pytest

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.calc import CalcDoc


def _tbl_data():
    vals = (
        ("Name", "Fruit", "Quantity"),
        ("Alice", "Apples", 3),
        ("Alice", "Oranges", 7),
        ("Bob", "Apples", 3),
        ("Alice", "Apples", 9),
        ("Bob", "Apples", 5),
        ("Bob", "Oranges", 6),
        ("Alice", "Oranges", 3),
        ("Alice", "Apples", 8),
        ("Alice", "Oranges", 1),
        ("Bob", "Oranges", 2),
        ("Bob", "Oranges", 7),
        ("Bob", "Apples", 1),
        ("Alice", "Apples", 8),
        ("Alice", "Oranges", 8),
        ("Alice", "Apples", 7),
        ("Bob", "Apples", 1),
        ("Bob", "Oranges", 9),
        ("Bob", "Oranges", 3),
        ("Alice", "Oranges", 4),
        ("Alice", "Apples", 9),
    )
    return vals


def test_cell_range_find_used(loader):
    doc = None
    try:
        doc = CalcDoc.create_doc(loader=loader)
        sheet = doc.sheets[0]
        vals = _tbl_data()

        sheet.set_array(values=vals, name="B2")

        rng = sheet.get_range(range_name="A1:H100")

        found_rng = rng.find_used_range()
        # B2:D22
        assert str(found_rng.range_obj) == "B2:D22"

    finally:
        if doc is not None:
            doc.close()


def test_cell_cursor_find_used(loader):
    doc = None
    try:
        doc = CalcDoc.create_doc(loader=loader)
        sheet = doc.sheets[0]
        vals = _tbl_data()

        sheet.set_array(values=vals, name="B2")

        rng = sheet.get_range(range_name="A1:H100")
        cursor = rng.create_cursor()

        found_rng = cursor.find_used_range_obj()
        # B2:D22
        assert str(found_rng) == "B2:D22"

    finally:
        if doc is not None:
            doc.close()
