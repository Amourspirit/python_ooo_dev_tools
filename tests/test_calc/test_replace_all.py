from __future__ import annotations
from typing import List, TYPE_CHECKING, cast
import pytest

if __name__ == "__main__":
    pytest.main([__file__])
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props
from ooodev.utils.uno_enum import UnoEnum
from ooodev.utils.table_helper import TableHelper
from ooodev.utils.color import CommonColor
from ooodev.office.calc import Calc

from com.sun.star.awt import FontWeight # const
from com.sun.star.table import XCellRange
from com.sun.star.util import XReplaceable
from com.sun.star.util import XSearchable
from com.sun.star.sheet import XSpreadsheet
from com.sun.star.sheet import XSpreadsheetDocument
from com.sun.star.util import XMergeable

if TYPE_CHECKING:
    from com.sun.star.table import CellHoriJustify as UnoCellHoriJustify
    from com.sun.star.table import CellVertJustify as UnoCellVertJustify


animals = ("ass", "cat", "cow", "cub", "doe", "dog", "elk", 
         "ewe", "fox", "gnu", "hog", "kid", "kit", "man",
         "orc", "pig", "pup", "ram", "rat", "roe", "sow", "yak")
names = ("Alf", "Amy", "Ivy", "Jet", "Ace", "Joe", "Joy",
         "Ben", "Gus", "Sky", "Bob", "Boo", "Bud", "Boy",
         "Sox", "Yin", "Kid", "Max", "Sly", "Bug", "Pat", "Tom")

def test_replace_all(loader) -> None:
    doc = Calc.create_doc(loader=loader)
    visible = False
    delay = 0 # 500
    if visible:
        GUI.set_visible(is_visible=visible, odoc=doc)
    sheet = Calc.get_sheet(doc=doc, index=0)

    try:
        total_rows = 15
        total_cols = 6

        def cb(row:int, col:int, prev) -> str:
            # return animals repeating until all cells are filled
            v = (row * total_cols) + col

            a_len = len(animals)
            if v > a_len - 1:
                i = (v % a_len)
            else:
                i = v
            return animals[i]
        
        arr = TableHelper.make_2d_array(num_rows=total_rows, num_cols=total_cols, val=cb)
        Calc.set_array(values=arr, sheet=sheet, name="A1")

        cell_range = Calc.get_cell_range(sheet=sheet, start_col=0, start_row=0, end_col=total_cols -1, end_row=total_rows-1) # A1:F15

        for w in ("cat", "cow", "dog", "kid", "ram", "ewe"):
            do_search(
                doc=doc,
                sheet=sheet,
                delay=delay,
                total_cols=total_cols,
                total_rows=total_rows,
                cell_range=cell_range,
                expected_len=5,
                search_wrd=w
            )
    finally:
        Lo.close(closeable=doc, deliver_ownership=False)

def do_search(doc: XSpreadsheetDocument, sheet: XSpreadsheet, delay:int, total_cols: int, total_rows: int, cell_range:XCellRange, expected_len: int, search_wrd: str) -> None:
        add_label(doc=doc, sheet=sheet, empty_row_num=total_rows+1, col_span=total_cols-1, text=f"Searching for {search_wrd}")
        ranges = search_iter(sheet=sheet, cell_range=cell_range, srch_str=search_wrd)
        len(ranges) == expected_len
        Lo.delay(delay)
        
        repl_str = names[animals.index(search_wrd)]
        add_label(doc=doc, sheet=sheet, empty_row_num=total_rows+1, col_span=total_cols-1, text=f"Replacing {search_wrd} with {repl_str}")
        replace_all(cell_range=cell_range, srch_str=search_wrd, repl_str=repl_str)
        for rng in ranges:
            replace_highlight(rng)
            cell = rng.getCellByPosition(0, 0)
            cstr = Calc.get_string(cell)
            assert cstr == repl_str
        Lo.delay(delay)

def search_iter(sheet: XSpreadsheet, cell_range: XCellRange, srch_str: str) -> List[XCellRange]:
    srch = Lo.qi(XSearchable, cell_range)
    sd = srch.createSearchDescriptor()
    
    sd.setSearchString(srch_str)
    sd.setPropertyValue("SearchWords", True)
                #   only complete words will be found
    
    cr = Lo.qi(XCellRange, srch.findFirst(sd))
    results = []
    if cr is not None:
        results.append(cr)
    while cr is not None:
        highlight(cr)
        cr = Lo.qi(XCellRange, srch.findNext(cr, sd))
        if cr is not None:
            results.append(cr)
    return results

def highlight(cr: XCellRange) -> None:
    Props.set_property(prop_set=cr, name="CharWeight", value=FontWeight.BOLD)
    Props.set_property(prop_set=cr, name="CharColor", value=CommonColor.DARK_BLUE)
    Props.set_property(prop_set=cr, name="CellBackColor", value=CommonColor.LIGHT_BLUE)


def replace_highlight(cr: XCellRange) -> None:
    Props.set_property(prop_set=cr, name="CharWeight", value=FontWeight.BOLD)
    Props.set_property(prop_set=cr, name="CharColor", value=CommonColor.INDIGO)
    Props.set_property(prop_set=cr, name="CellBackColor", value=0x99D700)

def search_all(sheet: XSpreadsheet, cell_range: XCellRange, srch_str: str) -> None:
    srch = Lo.qi(XSearchable, cell_range)
    sd = srch.createSearchDescriptor()
    
    sd.setSearchString(srch_str)
    sd.setPropertyValue("SearchWords", True)
    
    match_crs = Calc.find_all(srch, sd)
    if match_crs is None:
        return
    
    for match in match_crs:
        highlight(match)
        

def replace_all(cell_range:XCellRange, srch_str: str, repl_str: str ) -> int:
    repl = Lo.qi(XReplaceable, cell_range)
    rd = repl.createReplaceDescriptor()
    rd.setSearchString(srch_str)
    rd.setReplaceString(repl_str)
    rd.setPropertyValue("SearchWords", True)
    # rd.setPropertyValue("SearchRegularExpression", True)
    
    count = repl.replaceAll(rd)
    return count

def add_label(doc: XSpreadsheetDocument, sheet: XSpreadsheet, empty_row_num: int, col_span: int, text:str) -> None:
    """
    Add a large text string to the first cell
    in the empty row. Make the cell bigger by merging a few cells, and taller
    The text is black and bold in a red cell, and is centered.
    """
    CellHoriJustify = cast('UnoCellHoriJustify', UnoEnum("com.sun.star.table.CellHoriJustify"))
    CellVertJustify = cast('UnoCellVertJustify', UnoEnum("com.sun.star.table.CellVertJustify"))
    Calc.goto_cell(cell_name=Calc.get_cell_str(col=0, row=empty_row_num), doc=doc)
    
    # Merge first few cells of the last row
    cell_range = Calc.get_cell_range(sheet=sheet, start_col=0, start_row=empty_row_num, end_col=col_span, end_row=empty_row_num)
    xmerge = Lo.qi(XMergeable, cell_range)
    xmerge.merge(True)
    
    # make the row taller
    Calc.set_row_height(sheet=sheet, height=18, idx=empty_row_num)
    cell = Calc.get_cell(sheet=sheet, col=0, row=empty_row_num)
    cell.setFormula(text)
    Props.set_property(prop_set=cell, name="CharWeight", value=FontWeight.BOLD)
    Props.set_property(prop_set=cell, name="CharHeight", value=24)
    Props.set_property(prop_set=cell, name="CellBackColor", value=CommonColor.MISTY_ROSE)
    Props.set_property(prop_set=cell, name="HoriJustify", value=CellHoriJustify.CENTER)
    Props.set_property(prop_set=cell, name="VertJustify", value=CellVertJustify.CENTER)