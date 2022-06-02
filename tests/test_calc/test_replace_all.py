from __future__ import annotations
from typing import List
import pytest

if __name__ == "__main__":
    pytest.main([__file__])
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props
from ooodev.utils.gen_util import TableHelper
from ooodev.utils.color import CommonColor
from ooodev.office.calc import Calc

from com.sun.star.awt import FontWeight # const
from com.sun.star.table import XCellRange
from com.sun.star.util import XReplaceable
from com.sun.star.util import XSearchable
from com.sun.star.sheet import XSpreadsheet

animals = ("ass", "cat", "cow", "cub", "doe", "dog", "elk", 
         "ewe", "fox", "gnu", "hog", "kid", "kit", "man",
         "orc", "pig", "pup", "ram", "rat", "roe", "sow", "yak")

def test_replace_all(loader) -> None:
    doc = Calc.create_doc(loader=loader)
    visible = True
    delay = 2000
    if visible:
        GUI.set_visible(is_visible=visible, odoc=doc)
    sheet = Calc.get_sheet(doc=doc, index=0)

    try:
        def cb(row:int, col:int, prev) -> str:
            if prev is None:
                return animals[0]
            r = row + 1
            c = col + 1
            v = (r * c) - 1
            
            if v > len(animals) - 1:
                i = (v % len(animals)) - 1
            else:
                i = v
            return animals[i]
        
        arr = TableHelper.make_2d_array(num_rows=15, num_cols=6, val=cb)
        Calc.set_array(values=arr, sheet=sheet, name="A1")
        Lo.delay(delay)
    finally:
        Lo.close(closeable=doc, deliver_ownership=False)


def search_iter(sheet: XSpreadsheet, cell_range: XCellRange, srch_str: str) -> int:
    srch = Lo.qi(XSearchable, cell_range)
    sd = srch.createSearchDescriptor()
    
    sd.setSearchString(srch_str)
    sd.setPropertyValue("SearchWords", True)
                #   only complete words will be found
    
    cr = Lo.qi(XCellRange, srch.findFirst(sd))
    count = 0
    while cr is not None:
        highlight(cr)
        cr = Lo.qi(XCellRange, srch.findNext(cr, sd))
        count += 1
    return count

def highlight(cr: XCellRange) -> None:
    Props.set_property(prop_set=cr, name="CharWeight", value=FontWeight.BOLD)
    Props.set_property(prop_set=cr, name="CharColor", value=CommonColor.DARK_BLUE)
    Props.set_property(prop_set=cr, name="CellBackColor", value=CommonColor.LIGHT_BLUE)


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