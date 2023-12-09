import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.utils.lo import Lo
from ooodev.write import Write, WriteDoc
from ooodev.utils.info import Info


def test_style_family(loader) -> None:
    doc = WriteDoc(Write.create_doc(loader))
    from ooodev.write import FamilyNamesKind
    from ooodev.write import StyleParaKind

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
        families = doc.get_style_families()
        assert families is not None

        family = families.get_style_family(FamilyNamesKind.PARAGRAPH_STYLES)
        assert family is not None
        assert family.has_by_name(StyleParaKind.HEADING_2.value)
        style = family.get_style(StyleParaKind.HEADING_2.value)
        assert style is not None
        assert style.get_name() == StyleParaKind.HEADING_2.value
    finally:
        doc.close_doc()


def test_character_style(loader) -> None:
    doc = WriteDoc(Write.create_doc(loader))
    try:
        # services
        # 'com.sun.star.style.Style'
        # 'com.sun.star.style.PageStyle'
        # 'com.sun.star.style.PageProperties'
        style = doc.get_style_character()
        assert style is not None
        assert style.get_name() == "Standard"

        from ooodev.write import StyleCharKind

        style = doc.get_style_character(StyleCharKind.DROP_CAPS)
        assert style is not None
        assert style.get_name() == StyleCharKind.DROP_CAPS.value

    finally:
        doc.close_doc()


def test_frame_style(loader) -> None:
    doc = WriteDoc(Write.create_doc(loader))
    try:
        style = doc.get_style_frame()
        assert style is not None
        assert style.get_name() == "Frame"

        from ooodev.write import StyleFrameKind

        style = doc.get_style_frame(StyleFrameKind.WATERMARK)
        assert style is not None
        assert style.get_name() == StyleFrameKind.WATERMARK.value
    finally:
        doc.close_doc()


def test_numbering_style(loader) -> None:
    doc = WriteDoc(Write.create_doc(loader))
    try:
        from ooodev.write import StyleListKind

        style = doc.get_style_numbering()
        assert style is not None
        assert style.get_name() == StyleListKind.NUM_123.value

        style = doc.get_style_numbering(StyleListKind.LIST_01)
        assert style is not None
        assert style.get_name() == StyleListKind.LIST_01.value

    finally:
        doc.close_doc()


def test_page_style(loader) -> None:
    doc = WriteDoc(Write.create_doc(loader))
    try:
        style = doc.get_style_page()
        assert style is not None
        assert style.get_name() == "Standard"

        from ooodev.write import WriterStylePageKind

        style = doc.get_style_page(WriterStylePageKind.FIRST_PAGE)
        assert style is not None
        assert style.get_name() == WriterStylePageKind.FIRST_PAGE.value
    finally:
        doc.close_doc()


def test_paragraph_style(loader) -> None:
    doc = WriteDoc(Write.create_doc(loader))
    try:
        style = doc.get_style_paragraph()
        assert style is not None
        assert style.get_name() == "Standard"

        from ooodev.write import StyleParaKind

        style = doc.get_style_paragraph(StyleParaKind.HEADING_2)
        assert style is not None
        assert style.get_name() == StyleParaKind.HEADING_2.value
    finally:
        doc.close_doc()


def test_table_style(loader) -> None:
    doc = WriteDoc(Write.create_doc(loader))
    try:
        style = doc.get_style_table()
        assert style is not None
        assert style.get_name() == "Default Style"

        style = doc.get_style_table("Academic")
        assert style is not None
        assert style.get_name() == "Academic"
    finally:
        doc.close_doc()
