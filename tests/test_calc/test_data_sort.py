from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.utils.props import Props
from ooodev.office.calc import Calc


def test_basic_sort(loader) -> None:
    # https://wiki.openoffice.org/wiki/Documentation/OOo3_User_Guides/Calc_Guide/Sorting
    #
    # The SortFields property requires an array, not a single field. Create this in Python by sending a tuple to uno.Any().
    # https://ask.libreoffice.org/t/using-the-sort-method-with-python/53535
    import uno
    from com.sun.star.util import XSortable
    from com.sun.star.table import TableSortField
    from com.sun.star.view import XSelectionSupplier

    visible = False
    delay = 0  # 2000

    def make_sort_asc_tbl(index: int, ascending: bool) -> TableSortField:
        sf = TableSortField()
        sf.Field = index
        sf.IsAscending = ascending
        sf.IsCaseSensitive = False
        return sf

    doc = Calc.create_doc(loader=loader)
    assert doc is not None, "Could not create new document"

    sheet = Calc.get_sheet(doc=doc, index=0)
    vals = ((1, 5, "One"), (4, 1, "Two"), (3, 1, "Three"), (7, 8, "Four"), (4, 2, "Five"))
    Calc.set_array(values=vals, sheet=sheet, name="A1")

    if visible:
        GUI.set_visible(is_visible=visible, odoc=doc)

    # 1. obtain an XSortable interface for the cell range
    source_range = Calc.get_cell_range(sheet=sheet, range_name="A1:C5")
    xsort = Lo.qi(XSortable, source_range)

    # Optinally Select the range to sort.
    # The only purpose would be to emphasize the sorted data.
    if visible:
        cursor = sheet.createCursorByRange(source_range)
        controller = Calc.get_controller(doc)
        sel = Lo.qi(XSelectionSupplier, controller)
        if sel:
            sel.select(cursor)

    # 2. specify the sorting criteria as a TableSortField array
    # The columns are numbered starting with 0, so
    # column A is 0, column B is 1, etc.
    # Sort column B (column 1) descending.
    sf1 = make_sort_asc_tbl(index=1, ascending=False)
    # If column B has two cells with the same value,
    # then use column A ascending to decide the order.
    sf2 = make_sort_asc_tbl(index=0, ascending=True)

    # 3. define a sort descriptor
    odesc = xsort.createSortDescriptor()
    odesc[3].Value = uno.Any("[]com.sun.star.table.TableSortField", (sf1, sf2))

    Lo.wait(delay)  # wait so user can see original before it is sorted
    source_range.sort(odesc)  # 4. do the sort
    arr = Calc.get_array(cell_range=source_range)
    Lo.wait(delay)
    Lo.close(closeable=doc, deliver_ownership=False)
    assert arr is not None
    assert arr[0][1] == 8  # was 5
    assert arr[4][1] == 1  # was 2


def test_data_sort(loader) -> None:
    # https://wiki.openoffice.org/wiki/Documentation/OOo3_User_Guides/Calc_Guide/Sorting

    import uno
    from com.sun.star.table import TableSortField
    from com.sun.star.util import XSortable
    from com.sun.star.view import XSelectionSupplier

    visible = False
    delay = 0  # 2000

    def make_sort_asc_tbl(index: int, ascending: bool) -> TableSortField:
        sf = TableSortField()
        sf.Field = index
        sf.IsAscending = ascending
        sf.IsCaseSensitive = False
        return sf

    doc = Calc.create_doc(loader=loader)
    assert doc is not None, "Could not create new document"
    if visible:
        GUI.set_visible(is_visible=visible, odoc=doc)
    sheet = Calc.get_sheet(doc=doc, index=0)
    vals = (
        ("Level", "Code", "No.", "Team", "Name"),
        ("BS", 20, 4, "B", "Elle"),
        ("BS", 20, 6, "C", "Sweet"),
        ("BS", 20, 2, "A", "Chcomic"),
        ("CS", 30, 5, "A", "Ally"),
        ("MS", 10, 1, "A", "Joker"),
        ("MS", 10, 3, "B", "Kevin"),
        ("CS", 30, 7, "C", "Tom"),
    )
    Calc.set_array(values=vals, sheet=sheet, name="A1:E8")  # or just "A1"

    # 1. obtain an XSortable interface for the cell range
    source_range = Calc.get_cell_range(sheet=sheet, range_name="A1:E8")
    xsort = Lo.qi(XSortable, source_range)

    # Optinally Select the range to sort.
    # The only purpose would be to emphasize the sorted data.
    if visible:
        cursor = sheet.createCursorByRange(source_range)
        controller = Calc.get_controller(doc)
        sel = Lo.qi(XSelectionSupplier, controller)
        if sel:
            sel.select(cursor)

    # 2. specify the sorting criteria as a TableSortField array
    sort_fields = (make_sort_asc_tbl(1, True), make_sort_asc_tbl(2, True))

    # 3. define a sort descriptor
    # props = Props.make_props(
    #     SortFields=uno.Any("[]com.sun.star.table.TableSortField", [*sort_fields]), ContainsHeader=True
    # )
    props = Props.make_props(SortFields=Props.any(*sort_fields), ContainsHeader=True)

    Lo.wait(delay)  # wait so user can see original before it is sorted
    # 4. do the sort
    xsort.sort(props)

    arr = Calc.get_array(cell_range=source_range)
    Lo.wait(delay)
    Lo.close(closeable=doc, deliver_ownership=False)
    assert arr is not None
    assert arr[0][4] == "Name"
    assert arr[1][4] == "Joker"
    assert arr[2][4] == "Kevin"
    assert arr[3][4] == "Chcomic"
    assert arr[4][4] == "Elle"
    assert arr[5][4] == "Sweet"
    assert arr[6][4] == "Ally"
    assert arr[7][4] == "Tom"
