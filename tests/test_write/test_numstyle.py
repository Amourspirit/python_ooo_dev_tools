import os
import pytest
from pathlib import Path
from typing import cast

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])

def test_numstyle(loader):
    from ooodev.utils.lo import Lo
    from ooodev.office.write import Write
    from ooodev.utils.gui import GUI
    from ooodev.exceptions import ex
    from functools import partial

    visible = True
    delay = 2000
    doc = Write.create_doc(loader)
    if visible:
        GUI.set_visible(visible, doc)

    cursor = Write.get_cursor(doc)
    append = partial(Write.append, cursor)
    para = partial(Write.append_para, cursor)
    nl = partial(Write.append_line, cursor)
    get_pos = partial(Write.get_position, cursor)


    Write.append_para(cursor, "The following points are important:")
    pos = Write.append_para(cursor, "Have a good breakfast")
    Write.append_para(cursor, "Have a good lunch")
    with pytest.raises(ex.PropertyNotFoundError):
        Write.style_prev_paragraph(cursor, "NumberingStyleName", "Numbering 1")
    Write.append_para(cursor, "Have a good dinner")
    nl()
    tvc = Write.get_view_cursor(doc)
    Write.style_left(tvc, 60, "NumberingStyleName", "Numbering 1")
    # numberingLevel
    # Write.style_left(cursor, pos, "NumberingIsNumber", True)
    # Write.style_left(cursor, pos, "NumberingLevel", 1)
    # "23250040901"
    # https://www.openoffice.org/api/docs/common/ref/com/sun/star/style/ParagraphProperties.html#NumberingStyleName


    Lo.delay(delay)
    Lo.close_doc(doc, False)
