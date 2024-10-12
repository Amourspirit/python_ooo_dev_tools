"""
Write 10 forumlae into a new Text doc and save as a pdf file.
"""

from typing import List
import pytest
import sys
from pathlib import Path

if __name__ == "__main__":
    pytest.main([__file__])
import uno

from ooodev.loader.lo import Lo
from ooodev.office.write import Write
from ooodev.gui.gui import GUI
from ooodev.utils.date_time_util import DateUtil
from ooodev.utils.info import Info
from ooodev.utils.props import Props

from com.sun.star.beans import XPropertySet
from com.sun.star.text import XTextDocument
from com.sun.star.style import XStyle
from com.sun.star.text import XTextCursor

from ooo.dyn.style.line_spacing import LineSpacing
from ooo.dyn.style.line_spacing_mode import LineSpacingMode

# on windows getting Fatal Python error: Aborted even though the test runs fine when run by itself.


@pytest.mark.skipif(sys.platform == "win32", reason="does not run on windows in a group")
def test_story_creator(loader, copy_fix_writer, tmp_path_fn):
    visible = True
    delay = 4_000
    test_doc = Path(copy_fix_writer("scandal.txt"))
    fnm_save = Path(tmp_path_fn, "bigStory.doc")
    Write.create_doc(loader)
    # could have captured doc as: doc = Write.open_doc(fnm=test_doc, loader=loader)
    # just confirming that Lo.XSCRIPTCONTEXT is working.
    # Document is not visibe at this point os it is not available via Lo.XSCRIPTCONTEXT.getDocument()
    # or Lo.this_component
    doc = Lo.lo_component
    assert doc is not None

    try:
        if visible:
            GUI.set_visible(visible, doc)
        create_para_style(doc, "adParagraph")

        xtext_range = doc.getText().getStart()
        Props.set_property(xtext_range, "ParaStyleName", "adParagraph")

        Write.set_header(text_doc=doc, text=f"From: {test_doc.name}")
        Write.set_a4_page_format(doc)
        Write.set_page_numbers(doc)

        cursor = Write.get_cursor(doc)

        read_text(fnm=test_doc, cursor=cursor)
        Write.end_paragraph(cursor)

        Write.append_para(cursor, f"Timestamp: {DateUtil.time_stamp()}")

        Lo.delay(delay)
        Write.save_doc(text_doc=doc, fnm=fnm_save)
        assert fnm_save.exists()
    finally:
        Lo.close_doc(doc, False)


def read_text(fnm: Path, cursor: XTextCursor) -> None:
    sb: List[str] = []
    with open(fnm, "r") as file:
        i = 0
        for ln in file:
            line = ln.rstrip()  # remove new line \n
            if len(line) == 0:
                if len(sb) > 0:
                    Write.append_para(cursor, " ".join(sb))
                sb.clear()
            elif line.startswith("Title: "):
                Write.append_para(cursor, line[7:])
                Write.style_prev_paragraph(cursor, "Title")
            elif line.startswith("Author: "):
                Write.append_para(cursor, line[8:])
                Write.style_prev_paragraph(cursor, "Subtitle")
            elif line.startswith("Part "):
                Write.append_para(cursor, line)
                Write.style_prev_paragraph(cursor, "Heading")
            else:
                sb.append(line)
            i += 1
            # if i > 20:
            #     break
        if len(sb) > 0:
            Write.append_para(cursor, " ".join(sb))


def create_para_style(doc: XTextDocument, style_name: str) -> None:
    para_styles = Info.get_style_container(doc=doc, family_style_name="ParagraphStyles")

    # create new paragraph style properties set
    para_style = Lo.create_instance_msf(XStyle, "com.sun.star.style.ParagraphStyle", raise_err=True)
    props = Lo.qi(XPropertySet, para_style, raise_err=True)

    # set some properties
    props.setPropertyValue("CharFontName", Info.get_font_general_name())
    props.setPropertyValue("CharHeight", 12.0)
    props.setPropertyValue("ParaBottomMargin", 400)  # 4mm, in 100th mm

    line_spacing = LineSpacing(Mode=LineSpacingMode.FIX, Height=600)
    props.setPropertyValue("ParaLineSpacing", line_spacing)

    para_styles.insertByName(style_name, props)

    # set paragraph line spacing to 6mm
