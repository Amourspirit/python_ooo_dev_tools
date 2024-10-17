from __future__ import annotations
import pytest
from pathlib import Path

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.calc import CalcDoc
from ooodev.utils.helper.dot_dict import DotDict


def test_sheet_custom_props(loader, tmp_path) -> None:
    # get_sheet is overload method.
    # testing each overload.
    doc = None
    try:
        pth = Path(tmp_path, "test_sheet_custom_props.ods")
        doc = CalcDoc.create_doc(loader)
        sheet = doc.sheets[0]
        dd = DotDict()
        dd.sheet_name = "Sheet One"
        dd.sheet_num = 1
        sheet.set_custom_properties(dd)

        dd = sheet.get_custom_properties()
        assert dd.sheet_name == "Sheet One"
        assert dd.sheet_num == 1

        sheet2 = doc.sheets.insert_sheet("Sheet Two")
        dd = DotDict()
        dd.sheet_name = "Sheet Two"
        dd.sheet_num = 2
        sheet2.set_custom_properties(dd)

        dd = sheet2.get_custom_properties()
        assert dd.sheet_name == "Sheet Two"
        assert dd.sheet_num == 2

        doc.save_doc(pth)
        doc.close()

        doc = CalcDoc.open_doc(pth)
        sheet = doc.sheets[0]
        dd = sheet.get_custom_properties()
        assert dd.sheet_name == "Sheet One"
        assert dd.sheet_num == 1

        sheet2 = doc.sheets[1]
        dd = sheet2.get_custom_properties()
        assert dd.sheet_name == "Sheet Two"
        assert dd.sheet_num == 2
        assert len(dd) == 2

    finally:
        if doc:
            doc.close()


def test_sheet_cell_custom_props(loader, tmp_path) -> None:
    # get_sheet is overload method.
    # testing each overload.
    doc = None
    try:
        pth = Path(tmp_path, "test_sheet_cell_custom_props.ods")
        doc = CalcDoc.create_doc(loader=loader)
        sheet = doc.sheets[0]
        cell = sheet["A1"]
        cell.set_custom_property("test1", "test_val1")
        val = cell.get_custom_property("test1")
        assert val == "test_val1"
        cell.set_custom_property("test2", 10)
        val = cell.get_custom_property("test2")
        assert val == 10

        dd = cell.get_custom_properties()
        assert dd.test1 == "test_val1"
        assert dd.test2 == 10

        dd.test3 = 20
        cell.set_custom_properties(dd)

        cell = sheet["B2"]
        dd = DotDict()
        dd.Ran = "Ran One"
        dd.StopMe = "Stop Me"
        dd.num = 100.3
        cell.set_custom_properties(dd)

        dd = cell.get_custom_properties()
        assert dd.Ran == "Ran One"
        assert dd.StopMe == "Stop Me"
        assert dd.num == 100.3

        doc.save_doc(pth)
        doc.close()

        doc = CalcDoc.open_doc(pth)
        sheet = doc.sheets[0]
        cell = sheet["A1"]
        dd = cell.get_custom_properties()
        assert dd.test1 == "test_val1"
        assert dd.test2 == 10
        assert dd.test3 == 20

        cell = sheet["B2"]
        dd = cell.get_custom_properties()
        dd = cell.get_custom_properties()
        assert dd.Ran == "Ran One"
        assert dd.StopMe == "Stop Me"
        assert dd.num == 100.3
        assert len(dd) == 3

    finally:
        if doc:
            doc.close()


def test_sheet_and_cell_custom_props(loader, tmp_path) -> None:
    # get_sheet is overload method.
    # testing each overload.
    doc = None
    try:

        pth = Path(tmp_path, "test_sheet_and_cell_custom_props.ods")
        doc = CalcDoc.create_doc(loader)
        sheet = doc.sheets[0]
        dd = DotDict()
        dd.sheet_name = "Sheet One"
        dd.sheet_num = 1
        sheet.set_custom_properties(dd)

        dd = sheet.get_custom_properties()
        assert dd.sheet_name == "Sheet One"
        assert dd.sheet_num == 1

        cell = sheet["A1"]
        cell.set_custom_property("test1", "test_val1")
        val = cell.get_custom_property("test1")
        assert val == "test_val1"
        cell.set_custom_property("test2", 10)
        val = cell.get_custom_property("test2")
        assert val == 10

        dd = cell.get_custom_properties()
        assert dd.test1 == "test_val1"
        assert dd.test2 == 10

        dd.test3 = 20
        cell.set_custom_properties(dd)

        cell = sheet["B2"]
        dd = DotDict()
        dd.Ran = "Ran One"
        dd.StopMe = "Stop Me"
        dd.num = 100.3
        cell.set_custom_properties(dd)

        dd = cell.get_custom_properties()
        assert dd.Ran == "Ran One"
        assert dd.StopMe == "Stop Me"
        assert dd.num == 100.3

        sheet2 = doc.sheets.insert_sheet("Sheet Two")
        dd = DotDict()
        dd.sheet_name = "Sheet Two"
        dd.sheet_num = 2
        sheet2.set_custom_properties(dd)

        dd = sheet2.get_custom_properties()
        assert dd.sheet_name == "Sheet Two"
        assert dd.sheet_num == 2

        cell = sheet2["A1"]
        cell.set_custom_property("sheet2_test1", "test_val1")
        val = cell.get_custom_property("sheet2_test1")
        assert val == "test_val1"
        cell.set_custom_property("sheet2_test2", 10)
        val = cell.get_custom_property("sheet2_test2")
        assert val == 10

        dd = cell.get_custom_properties()
        assert dd.sheet2_test1 == "test_val1"
        assert dd.sheet2_test2 == 10

        dd.test3 = 20
        cell.set_custom_properties(dd)

        cell = sheet2["B2"]
        dd = DotDict()
        dd.Ran = "Ran One"
        dd.StopMe = "Stop Me"
        dd.num = 100.3
        cell.set_custom_properties(dd)

        dd = cell.get_custom_properties()
        assert dd.Ran == "Ran One"
        assert dd.StopMe == "Stop Me"
        assert dd.num == 100.3

        # test cell moved.
        sheet2.insert_row(1, 1)
        sheet2.insert_column(1, 1)
        cell.refresh()
        dd = cell.get_custom_properties()
        assert dd.Ran == "Ran One"
        assert dd.StopMe == "Stop Me"
        assert dd.num == 100.3

        doc.save_doc(pth)
        doc.close()

        doc = CalcDoc.open_doc(pth)
        sheet = doc.sheets[0]
        dd = sheet.get_custom_properties()
        assert dd.sheet_name == "Sheet One"
        assert dd.sheet_num == 1

        cell = sheet["A1"]
        dd = cell.get_custom_properties()
        assert dd.test1 == "test_val1"
        assert dd.test2 == 10
        assert dd.test3 == 20

        cell = sheet["B2"]
        dd = cell.get_custom_properties()
        dd = cell.get_custom_properties()
        assert dd.Ran == "Ran One"
        assert dd.StopMe == "Stop Me"
        assert dd.num == 100.3
        assert len(dd) == 3

        sheet2 = doc.sheets[1]
        dd = sheet2.get_custom_properties()
        assert dd.sheet_name == "Sheet Two"
        assert dd.sheet_num == 2
        assert len(dd) == 2

        cell = sheet2["A1"]
        val = cell.get_custom_property("sheet2_test1")
        assert val == "test_val1"
        val = cell.get_custom_property("sheet2_test2")
        assert val == 10

        dd = cell.get_custom_properties()
        assert dd.sheet2_test1 == "test_val1"
        assert dd.sheet2_test2 == 10

        # cell was moved down and right one row.
        cell = sheet2["C3"]
        dd = cell.get_custom_properties()
        assert dd.Ran == "Ran One"
        assert dd.StopMe == "Stop Me"
        assert dd.num == 100.3

    finally:
        if doc:
            doc.close()


def test_sheet_and_cell_custom_props_renamed_sheet(loader, tmp_path) -> None:
    # get_sheet is overload method.
    # testing each overload.
    doc = None
    try:

        pth = Path(tmp_path, "test_sheet_and_cell_custom_props_renamed_sheet.ods")
        doc = CalcDoc.create_doc(loader)
        sheet = doc.sheets[0]
        dd = DotDict()
        dd.sheet_name = "Sheet One"
        dd.sheet_num = 1
        sheet.set_custom_properties(dd)

        dd = sheet.get_custom_properties()
        assert dd.sheet_name == "Sheet One"
        assert dd.sheet_num == 1

        cell = sheet["A1"]
        cell.set_custom_property("test1", "test_val1")
        val = cell.get_custom_property("test1")
        assert val == "test_val1"
        cell.set_custom_property("test2", 10)
        val = cell.get_custom_property("test2")
        assert val == 10

        dd = cell.get_custom_properties()
        assert dd.test1 == "test_val1"
        assert dd.test2 == 10

        dd.test3 = 20
        cell.set_custom_properties(dd)

        cell = sheet["B2"]
        dd = DotDict()
        dd.Ran = "Ran One"
        dd.StopMe = "Stop Me"
        dd.num = 100.3
        cell.set_custom_properties(dd)

        dd = cell.get_custom_properties()
        assert dd.Ran == "Ran One"
        assert dd.StopMe == "Stop Me"
        assert dd.num == 100.3

        sheet2 = doc.sheets.insert_sheet("Sheet Two")
        dd = DotDict()
        dd.sheet_name = "Sheet Two"
        dd.sheet_num = 2
        sheet2.set_custom_properties(dd)

        dd = sheet2.get_custom_properties()
        assert dd.sheet_name == "Sheet Two"
        assert dd.sheet_num == 2

        cell = sheet2["A1"]
        cell.set_custom_property("sheet2_test1", "test_val1")
        val = cell.get_custom_property("sheet2_test1")
        assert val == "test_val1"
        cell.set_custom_property("sheet2_test2", 10)
        val = cell.get_custom_property("sheet2_test2")
        assert val == 10

        dd = cell.get_custom_properties()
        assert dd.sheet2_test1 == "test_val1"
        assert dd.sheet2_test2 == 10

        dd.test3 = 20
        cell.set_custom_properties(dd)

        cell = sheet2["B2"]
        dd = DotDict()
        dd.Ran = "Ran One"
        dd.StopMe = "Stop Me"
        dd.num = 100.3
        cell.set_custom_properties(dd)

        dd = cell.get_custom_properties()
        assert dd.Ran == "Ran One"
        assert dd.StopMe == "Stop Me"
        assert dd.num == 100.3

        # test cell moved.
        sheet2.insert_row(1, 1)
        cell.refresh()
        dd = cell.get_custom_properties()
        assert dd.Ran == "Ran One"
        assert dd.StopMe == "Stop Me"
        assert dd.num == 100.3

        sheet = doc.sheets[0]
        sheet.name = "Sheet One Renamed"
        sheet2 = doc.sheets[1]
        sheet2.name = "Sheet Two Renamed"

        doc.save_doc(pth)
        doc.close()

        doc = CalcDoc.open_doc(pth)
        sheet = doc.sheets[0]
        dd = sheet.get_custom_properties()
        assert dd.sheet_name == "Sheet One"
        assert dd.sheet_num == 1

        cell = sheet["A1"]
        dd = cell.get_custom_properties()
        assert dd.test1 == "test_val1"
        assert dd.test2 == 10
        assert dd.test3 == 20

        cell = sheet["B2"]
        dd = cell.get_custom_properties()
        dd = cell.get_custom_properties()
        assert dd.Ran == "Ran One"
        assert dd.StopMe == "Stop Me"
        assert dd.num == 100.3
        assert len(dd) == 3

        sheet2 = doc.sheets[1]
        dd = sheet2.get_custom_properties()
        assert dd.sheet_name == "Sheet Two"
        assert dd.sheet_num == 2
        assert len(dd) == 2

        cell = sheet2["A1"]
        val = cell.get_custom_property("sheet2_test1")
        assert val == "test_val1"
        val = cell.get_custom_property("sheet2_test2")
        assert val == 10

        dd = cell.get_custom_properties()
        assert dd.sheet2_test1 == "test_val1"
        assert dd.sheet2_test2 == 10

        # cell was moved down one row.
        cell = sheet2["B3"]
        dd = cell.get_custom_properties()
        assert dd.Ran == "Ran One"
        assert dd.StopMe == "Stop Me"
        assert dd.num == 100.3

    finally:
        if doc:
            doc.close()


def test_sheet_custom_cell(loader, tmp_path) -> None:
    # get_sheet is overload method.
    # testing each overload.
    doc = None
    try:

        doc = CalcDoc.create_doc(loader)
        sheet = doc.sheets[0]

        for i in range(1, 5):
            cell = sheet[f"A{i}"]
            dd = DotDict()
            dd.cell_name = f"A{i}"
            dd.cell_num = i
            cell.set_custom_properties(dd)

        for i in range(1, 5):
            cell = sheet[f"C{i}"]
            dd = DotDict()
            dd.cell_name = f"C{i}"
            dd.cell_num = i
            cell.set_custom_properties(dd)

        for i in range(1, 5):
            cell = sheet[f"D{i}"]
            dd = DotDict()
            dd.cell_name = f"d{i}"
            dd.cell_num = i
            dd.my_prop = f"my prop{i}"
            cell.set_custom_properties(dd)

        cell_props = sheet.custom_cell_properties.get_cell_properties()
        assert len(cell_props) == 12

        for cell_obj, props_set in cell_props.items():
            cell = sheet[cell_obj]
            assert "cell_name" in props_set
            assert "cell_num" in props_set
            for prop_name in props_set:
                assert cell.has_custom_property(prop_name)

        cell_props = sheet.custom_cell_properties.get_cell_properties("my_prop")
        assert len(cell_props) == 4

        for cell_obj, props_set in cell_props.items():
            cell = sheet[cell_obj]
            assert "cell_name" in props_set
            assert "cell_num" in props_set
            assert "my_prop" in props_set
            for prop_name in props_set:
                assert cell.has_custom_property(prop_name)

        sheet.custom_cell_properties.remove_all_custom_properties()
        assert sheet.draw_page.forms.has_by_name(sheet.custom_cell_properties.form_name) is False
        cell_props = sheet.custom_cell_properties.get_cell_properties()
        assert len(cell_props) == 0

    finally:
        if doc:
            doc.close()


def test_doc_custom_props(loader, tmp_path) -> None:
    # get_sheet is overload method.
    # testing each overload.
    doc = None
    try:
        pth = Path(tmp_path, "test_doc_custom_props.ods")
        doc = CalcDoc.create_doc(loader)

        dd = doc.get_custom_properties()
        assert len(dd) == 0

        dd = DotDict()
        dd.test1 = 10
        dd.test2 = "Hello World"
        dd.test3 = None
        doc.set_custom_properties(dd)
        test1 = doc.get_custom_property("test1")
        assert test1 == 10
        test2 = doc.get_custom_property("test2")
        assert test2 == "Hello World"
        test3 = doc.get_custom_property("test3")
        assert test3 is None

        doc.save_doc(pth)
        doc.close()

        doc = CalcDoc.open_doc(pth)
        dd = doc.get_custom_properties()
        assert dd.test1 == 10
        assert dd.test2 == "Hello World"
        assert dd.test3 is None

        test1 = doc.get_custom_property("test1")
        assert test1 == 10
        test2 = doc.get_custom_property("test2")
        assert test2 == "Hello World"
        test3 = doc.get_custom_property("test3")
        assert test3 is None

    finally:
        if doc:
            doc.close()
