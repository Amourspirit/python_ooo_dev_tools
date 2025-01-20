from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])


def test_named_ranges_sheet(copy_fix_calc, loader) -> None:
    # Test adding controls to a cell and a range
    # The controls are found when calling "cell.control.current_control" by shape position.
    from ooodev.calc import CalcDoc

    # from ooo.dyn.sheet.named_range_flag import NamedRangeFlagEnum
    from com.sun.star.table import CellAddress

    doc_path = copy_fix_calc("small_totals.ods")
    # Sheet Marks, 0 - cell_a2_sheet $Marks.$A$3
    # Sheet Marks, 0 - details_sheet $Marks.$A$2:$E$7
    # Global - details_mark2 $Marks2.$A$2:$E$7
    # Global - mark2_a2 $Marks2.$A$2
    doc = None
    try:
        doc = CalcDoc.open_doc(fnm=doc_path, loader=loader)
        sheet = doc.sheets[0]

        assert sheet.named_ranges.has_by_name("cell_a2_sheet")
        nr = sheet.named_ranges.get_by_name("cell_a2_sheet")
        assert nr.get_content() == "$Marks.$A$3"

        assert sheet.named_ranges.has_by_name("details_sheet")
        nr = sheet.named_ranges.get_by_name("details_sheet")
        assert nr.get_content() == "$Marks.$A$2:$E$7"

        assert sheet.named_ranges.has_by_name("details_mark2") is False

        assert sheet.named_ranges.has_by_name("non_existing") is False

        ca = CellAddress()
        ca.Sheet = 0
        ca.Column = 0
        ca.Row = 1
        sheet.named_ranges.add_new_by_name(
            name="new_range",
            content="$Marks.$A$2:$D$7",
            position=ca,
        )
        assert sheet.named_ranges.has_by_name("new_range")
        nr = sheet.named_ranges.get_by_name("new_range")
        assert nr.get_content() == "$Marks.$A$2:$D$7"

    finally:
        if doc is not None:
            doc.close()


def test_get_named_range_sheet(copy_fix_calc, loader) -> None:
    # Test adding controls to a cell and a range
    # The controls are found when calling "cell.control.current_control" by shape position.
    from ooodev.calc import CalcDoc
    from ooodev.exceptions import ex as mEx  # noqa: F401

    doc_path = copy_fix_calc("small_totals.ods")
    # Sheet Marks, 0 - cell_a2_sheet $Marks.$A$3
    # Sheet Marks, 0 - details_sheet $Marks.$A$2:$E$7
    # Global - details_mark2 $Marks2.$A$2:$E$7
    # Global - mark2_a2 $Marks2.$A$2
    doc = None
    try:
        doc = CalcDoc.open_doc(fnm=doc_path, loader=loader)
        sheet = doc.sheets[0]

        rng = sheet.get_range(named_range="cell_a2_sheet")
        assert rng is not None
        assert str(rng.range_obj) == "A3:A3"

        rng = sheet.get_range(named_range="details_sheet")
        assert rng is not None
        assert str(rng.range_obj) == "A2:E7"

        with pytest.raises(mEx.MissingNameError):
            sheet.get_range(named_range="details_mark2")

        with pytest.raises(mEx.MissingNameError):
            sheet.get_range(named_range="non_existing")

    finally:
        if doc is not None:
            doc.close()


def test_get_named_cursor_sheet(copy_fix_calc, loader) -> None:
    # Test adding controls to a cell and a range
    # The controls are found when calling "cell.control.current_control" by shape position.
    from ooodev.calc import CalcDoc
    from ooodev.exceptions import ex as mEx  # noqa: F401

    doc_path = copy_fix_calc("small_totals.ods")
    # Sheet Marks, 0 - cell_a2_sheet $Marks.$A$3
    # Sheet Marks, 0 - details_sheet $Marks.$A$2:$E$7
    # Global - details_mark2 $Marks2.$A$2:$E$7
    # Global - mark2_a2 $Marks2.$A$2
    doc = None
    try:
        doc = CalcDoc.open_doc(fnm=doc_path, loader=loader)
        sheet = doc.sheets[0]

        cursor = sheet.create_cursor_by_range(named_range="cell_a2_sheet")
        assert cursor is not None

        cursor = sheet.create_cursor_by_range(named_range="details_sheet")
        assert cursor is not None

        with pytest.raises(mEx.MissingNameError):
            sheet.create_cursor_by_range(named_range="details_mark2")

        with pytest.raises(mEx.MissingNameError):
            sheet.create_cursor_by_range(named_range="non_existing")

    finally:
        if doc is not None:
            doc.close()
