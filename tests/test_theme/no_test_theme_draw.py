from __future__ import annotations
import pytest
from os import getenv

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.theme import ThemeDraw, ThemeKind
from ooodev.utils.info import Info


def test_theme(loader) -> None:
    ver = Info.version_info
    if ver < (7, 5, 0, 0):
        return

    theme = ThemeDraw(theme_name=ThemeKind.LIBRE_OFFICE)
    assert theme.grid_color >= -1
    assert isinstance(theme.grid_visible, bool)


def test_theme_default(loader) -> None:
    ver = Info.version_info
    if ver < (7, 5, 0, 0):
        return

    theme = ThemeDraw()
    assert theme.grid_color >= -1
    assert isinstance(theme.grid_visible, bool)


def test_theme_dark(loader) -> None:
    ver = Info.version_info
    if ver < (7, 5, 0, 0):
        return

    theme = ThemeDraw(theme_name=ThemeKind.LIBRE_OFFICE_DARK)
    assert theme.grid_color >= -1
    assert isinstance(theme.grid_visible, bool)
