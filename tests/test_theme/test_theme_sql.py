from __future__ import annotations
import pytest
from os import getenv

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.theme import ThemeSql, ThemeKind
from ooodev.utils.info import Info


def test_theme(loader) -> None:
    ver = Info.version_info
    if ver < (7, 5, 0, 0):
        return

    theme = ThemeSql(theme_name=ThemeKind.LIBRE_OFFICE)
    assert theme.comment_color >= -1
    assert theme.identifier_color >= -1
    assert theme.keyword_color >= -1
    assert theme.number_color >= -1
    assert theme.operator_color >= -1
    assert theme.parameter_color >= -1
    assert theme.string_color >= -1


def test_theme_default(loader) -> None:
    ver = Info.version_info
    if ver < (7, 5, 0, 0):
        return

    theme = ThemeSql()
    assert theme.comment_color >= -1
    assert theme.identifier_color >= -1
    assert theme.keyword_color >= -1
    assert theme.number_color >= -1
    assert theme.operator_color >= -1
    assert theme.parameter_color >= -1
    assert theme.string_color >= -1


def test_theme_dark(loader) -> None:
    ver = Info.version_info
    if ver < (7, 5, 0, 0):
        return

    theme = ThemeSql(theme_name=ThemeKind.LIBRE_OFFICE_DARK)
    assert theme.comment_color >= -1
    assert theme.identifier_color >= -1
    assert theme.keyword_color >= -1
    assert theme.number_color >= -1
    assert theme.operator_color >= -1
    assert theme.parameter_color >= -1
    assert theme.string_color >= -1
