from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.theme import ThemeTextDoc
from ooodev.utils.info import Info
from ooodev.write import WriteDoc


def test_theme_default(loader) -> None:
    ver = Info.version_info
    if ver < (7, 6, 0, 0):
        return
    doc = None
    try:
        doc = WriteDoc.create_doc()
        theme = ThemeTextDoc()
        assert theme.grammar_color >= 0
        assert theme.direct_cursor_color >= 0
        assert isinstance(theme.direct_cursor_visible, bool)
        assert theme.field_shadings_color >= 0
        assert theme.header_footer_mark_color >= 0
        assert theme.index_table_shadings_color >= 0
        assert isinstance(theme.index_table_shadings_visible, bool)
        assert theme.page_columns_breaks_color >= 0
        assert theme.script_indicator_color >= 0
        assert theme.section_boundaries_color >= 0
        assert isinstance(theme.section_boundaries_visible, bool)
        assert theme.grid_color >= 0
    finally:
        if doc:
            doc.close()


def test_theme_no_exist(loader) -> None:
    ver = Info.version_info
    if ver < (7, 6, 0, 0):
        return

    doc = None
    try:
        doc = WriteDoc.create_doc()
        with pytest.raises(ValueError):
            _ = ThemeTextDoc(theme_name="some random name that does not exist")
    finally:
        if doc:
            doc.close()
