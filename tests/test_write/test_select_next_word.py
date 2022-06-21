from typing import TYPE_CHECKING, cast
import pytest
from pathlib import Path
# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.utils.lo import Lo
from ooodev.utils.info import Info
from ooodev.office.write import Write
from ooodev.utils.gui import GUI
from ooodev.utils.selection import WordTypeEnum

from com.sun.star.text import XTextDocument
from com.sun.star.text import XTextRange



def test_select_next_word(loader, copy_fix_writer):
    
    visible = True
    delay = 2000

    # text file opens with each new line being considered a paragraph break.
    test_doc = copy_fix_writer("scandalStart.odt")
    doc = Write.open_doc(test_doc, loader)
    try:
        if visible:
            GUI.set_visible(visible, doc)
       
        assert Write.is_anything_selected(doc) == False
        Write.select_next_word(doc)
        assert Write.is_anything_selected(doc)
        
        Lo.delay(delay)
    finally:
        Lo.close_doc(doc, False)