import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.loader.lo import Lo
from ooodev.office.write import Write
from ooodev.utils.info import Info
from ooodev.utils.gui import GUI
from ooodev.utils.date_time_util import DateUtil


def test_style_family(loader) -> None:
    doc = Write.create_doc(loader)
    try:
        # 'com.sun.star.style.StyleFamilies'
        # styles contains a collection of com.sun.star.style.StyleFamily
        # No. of Style Family Names: 7
        # CellStyles - StyleFamily
        # CharacterStyles - StyleFamily
        # FrameStyles - StyleFamily
        # NumberingStyles - StyleFamily
        # PageStyles - StyleFamily
        # ParagraphStyles - StyleFamily
        # TableStyles - StyleFamily

        # a style family contains a collection of com.sun.star.style.XStyle
        styles = Info.get_style_families(doc)
        assert styles is not None
    finally:
        Lo.close_doc(doc, False)


def test_page_style(loader) -> None:
    doc = Write.create_doc(loader)
    try:
        # services
        # 'com.sun.star.style.Style'
        # 'com.sun.star.style.PageStyle'
        # 'com.sun.star.style.PageProperties'
        style = Info.get_page_style_props(doc)
        assert style is not None
    finally:
        Lo.close_doc(doc, False)


def test_paragraph_style(loader) -> None:
    doc = Write.create_doc(loader)
    try:
        # services
        # 'com.sun.star.style.Style'
        # 'com.sun.star.style.PageStyle'
        # 'com.sun.star.style.PageProperties'
        style = Info.get_paragraph_style_props(doc)
        assert style is not None
    finally:
        Lo.close_doc(doc, False)
