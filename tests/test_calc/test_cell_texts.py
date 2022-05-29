from __future__ import annotations
import pytest
if __name__ == "__main__":
    pytest.main([__file__])
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.utils.color import CommonColor
from ooodev.office.write import Write
from ooodev.utils.props import Props
from ooodev.office.calc import Calc

def test_paragraph(loader) -> None:
    from com.sun.star.text import XText
    doc = Calc.create_doc(loader=loader)
    assert doc is not None, "Could not create new document"
    visible = False
    delay = 0
    if visible:
        GUI.set_visible(is_visible=visible, odoc=doc)
    sheet = Calc.get_sheet(doc=doc, index=0)
    
    Calc.highlight_range(sheet=sheet, range_name="A2:C7", headline="Cells and Cell Ranges")
    Calc.set_col_width(sheet=sheet, width=80, idx=1)
    xcell = Calc.get_cell(sheet=sheet, cell_name="B4")
    xtext = Lo.qi(XText, xcell)
    cursor = xtext.createTextCursor()
    # Insert two text paragraphs and a hyperlink into the cell
    Write.append_para(cursor=cursor, text="Text in first line.")
    Write.append(cursor=cursor, text="And a ")
    Write.add_hyperlink(cursor=cursor, label="hyperlink", url_str="https://github.com/Amourspirit/python_ooo_dev_tools")

    # beautify the cell
    # properties from styles.CharacterProperties
    Props.set_property(prop_set=xcell, name="CharColor", value=CommonColor.DARK_BLUE)
    Props.set_property(prop_set=xcell, name="CharHeight", value=18.0)
    
    # property from styles.ParagraphProperties
    Props.set_property(prop_set=xcell, name="ParaLeftMargin", value=500)
    Calc.add_annotation(sheet=sheet, cell_name="B4", msg="This annotation is located at B4")

    para_text = 'Text in first line.\nAnd a hyperlink'
    text = Calc.get_string(sheet=sheet, cell_name='B4')
    assert text == para_text
    assert xtext.getString() == para_text
    
    Lo.delay(delay)
    Lo.close(closeable=doc, deliver_ownership=False)