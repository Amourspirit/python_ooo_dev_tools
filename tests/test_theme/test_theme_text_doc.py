from __future__ import annotations
import pytest
from os import getenv

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.theme import ThemeTextDoc, ThemeKind
from ooodev.utils.info import Info


def test_theme(loader) -> None:
    if getenv("DEV_CONTAINER"):
        pytest.skip("Skip test in container: May be no theme data")
    ver = Info.version_info
    if ver < (7, 5, 0, 0):
        return

    theme = ThemeTextDoc(theme_name=ThemeKind.LIBRE_OFFICE)
    assert theme.boundaries_color >= -1
    assert isinstance(theme.boundaries_visible, bool)
    assert theme.doc_color >= -1
    assert theme.grammar_color >= -1
    assert theme.direct_cursor_color >= -1
    assert isinstance(theme.direct_cursor_visible, bool)
    assert theme.field_shadings_color >= -1
    assert theme.header_footer_mark_color >= -1
    assert theme.index_table_shadings_color >= -1
    assert isinstance(theme.index_table_shadings_visible, bool)
    assert theme.page_columns_breaks_color >= -1
    assert theme.script_indicator_color >= -1
    assert theme.section_boundaries_color >= -1
    assert isinstance(theme.section_boundaries_visible, bool)
    assert theme.grid_color >= -1


def test_theme_default(loader) -> None:
    if getenv("DEV_CONTAINER"):
        pytest.skip("Skip test in container: May be no theme data")
    ver = Info.version_info
    if ver < (7, 5, 0, 0):
        return

    theme = ThemeTextDoc()
    assert theme.boundaries_color >= -1
    assert isinstance(theme.boundaries_visible, bool)
    assert theme.doc_color >= -1
    assert theme.grammar_color >= -1
    assert theme.direct_cursor_color >= -1
    assert isinstance(theme.direct_cursor_visible, bool)
    assert theme.field_shadings_color >= -1
    assert theme.header_footer_mark_color >= -1
    assert theme.index_table_shadings_color >= -1
    assert isinstance(theme.index_table_shadings_visible, bool)
    assert theme.page_columns_breaks_color >= -1
    assert theme.script_indicator_color >= -1
    assert theme.section_boundaries_color >= -1
    assert isinstance(theme.section_boundaries_visible, bool)
    assert theme.grid_color >= -1


def test_theme_dark(loader) -> None:
    if getenv("DEV_CONTAINER"):
        pytest.skip("Skip test in container: May be no theme data")
    ver = Info.version_info
    if ver < (7, 5, 0, 0):
        return

    theme = ThemeTextDoc(theme_name=ThemeKind.LIBRE_OFFICE_DARK)
    assert theme.boundaries_color >= -1
    assert isinstance(theme.boundaries_visible, bool)
    assert theme.doc_color >= -1
    assert theme.grammar_color >= -1
    assert theme.direct_cursor_color >= -1
    assert isinstance(theme.direct_cursor_visible, bool)
    assert theme.field_shadings_color >= -1
    assert theme.header_footer_mark_color >= -1
    assert theme.index_table_shadings_color >= -1
    assert isinstance(theme.index_table_shadings_visible, bool)
    assert theme.page_columns_breaks_color >= -1
    assert theme.script_indicator_color >= -1
    assert theme.section_boundaries_color >= -1
    assert isinstance(theme.section_boundaries_visible, bool)
    assert theme.grid_color >= -1
