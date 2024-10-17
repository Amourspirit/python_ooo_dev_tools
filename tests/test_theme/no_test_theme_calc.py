from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.theme import ThemeCalc, ThemeKind
from ooodev.utils.info import Info


def test_theme(loader) -> None:
    ver = Info.version_info
    if ver < (7, 5, 0, 0):
        return

    theme = ThemeCalc(theme_name=ThemeKind.LIBRE_OFFICE)
    assert theme.detective_color >= -1
    assert theme.detective_error_color >= -1
    assert theme.formula_color >= -1
    assert theme.hidden_col_row_color >= -1
    assert isinstance(theme.hidden_col_row_visible, bool)
    assert theme.notes_background_color >= -1
    assert theme.grid_color >= -1
    assert theme.page_break_auto_color >= -1
    assert theme.page_break_manual_color >= -1
    assert theme.text_color >= -1
    assert theme.value_color >= -1


def test_theme_default(loader) -> None:
    ver = Info.version_info
    if ver < (7, 5, 0, 0):
        return

    theme = ThemeCalc()
    assert theme.detective_color >= -1
    assert theme.detective_error_color >= -1
    assert theme.formula_color >= -1
    assert theme.hidden_col_row_color >= -1
    assert isinstance(theme.hidden_col_row_visible, bool)
    assert theme.notes_background_color >= -1
    assert theme.grid_color >= -1
    assert theme.page_break_auto_color >= -1
    assert theme.page_break_manual_color >= -1
    assert theme.text_color >= -1
    assert theme.value_color >= -1


def test_theme_dark(loader) -> None:
    ver = Info.version_info
    if ver < (7, 5, 0, 0):
        return

    theme = ThemeCalc(theme_name=ThemeKind.LIBRE_OFFICE_DARK)
    assert theme.detective_color >= -1
    assert theme.detective_error_color >= -1
    assert theme.formula_color >= -1
    assert theme.hidden_col_row_color >= -1
    assert isinstance(theme.hidden_col_row_visible, bool)
    assert theme.notes_background_color >= -1
    assert theme.grid_color >= -1
    assert theme.page_break_auto_color >= -1
    assert theme.page_break_manual_color >= -1
    assert theme.text_color >= -1
    assert theme.value_color >= -1
