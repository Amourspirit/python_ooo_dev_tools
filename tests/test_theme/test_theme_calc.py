from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.theme import ThemeCalc, ThemeKind
from ooodev.utils.info import Info


def test_theme(loader) -> None:
    ver = Info.version_info
    if ver < (7, 6, 0, 0):
        return

    theme = ThemeCalc(theme_name=ThemeKind.AUTOMATIC)
    assert theme.detective_color >= 0
    assert theme.detective_error_color >= 0
    assert theme.formula_color >= 0
    assert theme.hidden_col_row_color >= 0
    assert isinstance(theme.hidden_col_row_visible, bool)
    assert theme.notes_background_color >= 0
    assert theme.grid_color >= 0
    assert theme.page_break_auto_color >= 0
    assert theme.page_break_manual_color >= 0
    assert theme.text_color >= 0
    assert theme.value_color >= 0


def test_theme_default(loader) -> None:
    ver = Info.version_info
    if ver < (7, 6, 0, 0):
        return

    theme = ThemeCalc()
    assert theme.detective_color >= 0
    assert theme.detective_error_color >= 0
    assert theme.formula_color >= 0
    assert theme.hidden_col_row_color >= 0
    assert isinstance(theme.hidden_col_row_visible, bool)
    assert theme.notes_background_color >= 0
    assert theme.grid_color >= 0
    assert theme.page_break_auto_color >= 0
    assert theme.page_break_manual_color >= 0
    assert theme.text_color >= 0
    assert theme.value_color >= 0


def test_theme_no_exist(loader) -> None:
    ver = Info.version_info
    if ver < (7, 6, 0, 0):
        return
    with pytest.raises(ValueError):
        _ = ThemeCalc(theme_name="some random name that does not exist")
