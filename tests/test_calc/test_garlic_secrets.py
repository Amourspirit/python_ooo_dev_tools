from __future__ import annotations
from typing import TYPE_CHECKING, cast
from pathlib import Path
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.utils.color import CommonColor
from ooodev.utils.props import Props
from ooodev.utils.view_state import ViewState
from ooodev.utils.uno_enum import UnoEnum
from ooodev.office.calc import Calc, GeneralFunction

from com.sun.star.awt import FontWeight
from com.sun.star.sheet import XCellRangesQuery
from com.sun.star.sheet import XSpreadsheet
from com.sun.star.sheet import XSpreadsheetDocument
from com.sun.star.util import XMergeable

if TYPE_CHECKING:
    from com.sun.star.table import CellHoriJustify as UnoCellHoriJustify
    from com.sun.star.table import CellVertJustify as UnoCellVertJustify
    from com.sun.star.table import CellContentType as UnoCellContentType

FNM = "produceSales.xlsx"
OUT_FNM = "garlicSecrets.ods"


def test_garlic_secrets(copy_fix_calc, loader, test_headless, capsys: pytest.CaptureFixture) -> None:
    doc_path: Path = copy_fix_calc(FNM)
    doc = Calc.open_doc(fnm=str(doc_path), loader=loader)
    if doc is None:
        Lo.close_office()
        assert False, f"Could not open {FNM}"
    visible = not test_headless
    delay = 500
    if visible:
        GUI.set_visible(is_visible=visible, odoc=doc)
    sheet = Calc.get_sheet(doc=doc, index=0)
    Calc.goto_cell(cell_name="A1", doc=doc)

    # freeze on row of view
    if visible:
        Calc.freeze_rows(doc=doc, num_rows=1)

    # find total for the "Total" column
    total_range = Calc.get_col_range(sheet=sheet, idx=3)
    total = Calc.compute_function(fn=GeneralFunction.SUM, cell_range=total_range)
    assert total == pytest.approx(231353.56999999937, rel=1e-5)

    increase_garlic_cost(doc=doc, sheet=sheet)  # takes several seconds

    # recalculate total
    total_range = Calc.get_col_range(sheet=sheet, idx=3)
    total = Calc.compute_function(fn=GeneralFunction.SUM, cell_range=total_range)
    assert total == pytest.approx(231488.57999999903, rel=1e-5)

    # add a label at the bottom of the data, and hide it
    empty_row_num = find_empty_row(sheet=sheet)
    assert empty_row_num == 5000  # zero based index
    add_garlic_label(doc=doc, sheet=sheet, empty_row_num=empty_row_num)
    # wait a bit before hiding last row
    Lo.delay(delay)
    row_range = Calc.get_row_range(sheet=sheet, idx=empty_row_num)
    Props.set_property(prop_set=row_range, name="IsVisible", value=False)

    if visible:
        Calc.unfreeze(doc=doc)

    # split window into 2 view panes
    cell_name = Calc.get_cell_str(col=0, row=empty_row_num - 2)
    # doesn't work with Calc.freeze()
    Calc.split_window(doc=doc, cell_name=cell_name)
    Lo.delay(delay)

    # access panes; make top pane show first row
    panes = Calc.get_view_panes(doc=doc)
    assert len(panes) > 0
    panes[0].setFirstVisibleRow(0)
    Lo.delay(delay)

    # display view properties
    ss_view = Calc.get_view(doc=doc)
    capsys.readouterr()  # clear buffer
    Props.show_obj_props(prop_kind="Spreadsheet view", obj=ss_view)
    cresult = capsys.readouterr()
    presult: str = cresult.out
    assert presult is not None
    ex_lst = presult.splitlines()
    assert len(ex_lst) == 30 or 31

    # show view data
    print(f"View Data: {Calc.get_view_data(doc=doc)}")
    cresult = capsys.readouterr()
    presult: str = cresult.out
    assert presult is not None
    ex_lst = presult.splitlines()
    assert len(ex_lst) == 1

    # show sheet states
    states = Calc.get_view_states(doc=doc)
    assert len(states) > 0
    str_state = ""
    for state in states:
        capsys.readouterr()  # clear buffer
        state.report()
        cresult = capsys.readouterr()
        presult: str = cresult.out
        ex_lst = presult.splitlines()
        assert len(ex_lst) == 7
        str_state = str(state)
        # [
        # 'Sheet View State',
        # "  Cursor pos (column, row): (0, 4998) or 'A4999'",
        # '  Sheet is split horizontally at 400',
        # '  Number of focused pane: 2',
        # '  Left column indicies of left/right panes: 0 / 0',
        # '  Top row indicies of upper/lower panes: 0 / 4998',
        # ''
        # ]

    # make top pane the active one in the first sheet
    states[0].move_pane_focus(dir=ViewState.PaneEnum.MOVE_UP)
    Calc.set_view_states(doc=doc, states=states)
    Calc.goto_cell(cell_name="A1", doc=doc)

    rev_states = Calc.get_view_states(doc=doc)
    assert len(rev_states) > 0
    for state in rev_states:
        capsys.readouterr()  # clear buffer
        state.report()
        cresult = capsys.readouterr()
        presult: str = cresult.out
        ex_lst = presult.splitlines()
        assert len(ex_lst) == 7

    rev_state = rev_states[0]
    str_rev_state = rev_state.to_string()
    assert str_rev_state != str_state
    orig_state = ViewState(str_state)

    assert orig_state.pane_focus_num == 2
    if not test_headless:
        assert rev_state.pane_focus_num == 0
    Lo.delay(delay)

    Calc.insert_row(sheet=sheet, idx=0)
    add_garlic_label(doc=doc, sheet=sheet, empty_row_num=0)
    lbl = Calc.get_string(sheet=sheet, cell_name="A1")
    assert lbl == "Top Secret Garlic Changes"

    save_path = Path(doc_path.parent, OUT_FNM)
    Lo.save_doc(doc=doc, fnm=str(save_path))
    assert save_path.exists()
    assert save_path.is_file()

    Lo.delay(delay)
    Lo.close(closeable=doc, deliver_ownership=False)


def increase_garlic_cost(doc: XSpreadsheetDocument, sheet: XSpreadsheet) -> int:
    """
    Iterate down the "Produce" column. If the text in the current cell is
    "Garlic" then change the corresponding "Cost Per Pound" cell by
    multiplying by 1.05, and changing its text to bold red.

    Return the "Produce" row index which is first empty.
    """

    CellContentType = cast("UnoCellContentType", UnoEnum("com.sun.star.table.CellContentType"))

    row = 0
    prod_cell = Calc.get_cell(sheet=sheet, col=0, row=row)  # produce column
    # iterate down produce column until an empty cell is reached
    while prod_cell.getType() != CellContentType.EMPTY:
        if prod_cell.getFormula() == "Garlic":
            # show the cell in-screen
            Calc.goto_cell(doc=doc, cell_name=Calc.get_cell_str(col=0, row=row))
            # change cost/pound column
            cost_cell = Calc.get_cell(sheet=sheet, col=1, row=row)
            cost_cell.setValue(1.05 * cost_cell.getValue())
            Props.set_property(prop_set=cost_cell, name="CharWeight", value=FontWeight.BOLD)
            Props.set_property(prop_set=cost_cell, name="CharColor", value=CommonColor.RED)
        row += 1
        prod_cell = Calc.get_cell(sheet=sheet, col=0, row=row)
    return row


def find_empty_row(sheet: XSpreadsheet) -> int:
    """
    Return the index of the first empty row by finding all the empty cell ranges in
    the first column, and return the smallest row index in those ranges.
    """

    # create a ranges query for the first column of the sheet
    cell_range = Calc.get_col_range(sheet=sheet, idx=0)
    Calc.print_address(cell_range=cell_range)
    cr_query = Lo.qi(XCellRangesQuery, cell_range)
    sc_ranges = cr_query.queryEmptyCells()
    addrs = sc_ranges.getRangeAddresses()
    Calc.print_addresses(*addrs)

    # find smallest row index
    row = -1
    if addrs is not None and len(addrs) > 0:
        row = addrs[0].StartRow
        for addr in addrs:
            if row < addr.StartRow:
                row = addr.StartRow
        print(f"First empty row is at position: {row}")
    else:
        print("Could not find an empty row")
    return row


def add_garlic_label(doc: XSpreadsheetDocument, sheet: XSpreadsheet, empty_row_num: int) -> None:
    """
    Add a large text string ("Top Secret Garlic Changes") to the first cell
    in the empty row. Make the cell bigger by merging a few cells, and taller
    The text is black and bold in a red cell, and is centered.
    """
    CellHoriJustify = cast("UnoCellHoriJustify", UnoEnum("com.sun.star.table.CellHoriJustify"))
    CellVertJustify = cast("UnoCellVertJustify", UnoEnum("com.sun.star.table.CellVertJustify"))
    Calc.goto_cell(cell_name=Calc.get_cell_str(col=0, row=empty_row_num), doc=doc)

    # Merge first few cells of the last row
    cell_range = Calc.get_cell_range(
        sheet=sheet, start_col=0, start_row=empty_row_num, end_col=3, end_row=empty_row_num
    )
    xmerge = Lo.qi(XMergeable, cell_range)
    xmerge.merge(True)

    # make the row taller
    Calc.set_row_height(sheet=sheet, height=18, idx=empty_row_num)
    cell = Calc.get_cell(sheet=sheet, col=0, row=empty_row_num)
    cell.setFormula("Top Secret Garlic Changes")
    Props.set_property(prop_set=cell, name="CharWeight", value=FontWeight.BOLD)
    Props.set_property(prop_set=cell, name="CharHeight", value=24)
    Props.set_property(prop_set=cell, name="CellBackColor", value=CommonColor.RED)
    Props.set_property(prop_set=cell, name="HoriJustify", value=CellHoriJustify.CENTER)
    Props.set_property(prop_set=cell, name="VertJustify", value=CellVertJustify.CENTER)
