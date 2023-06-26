from __future__ import annotations
import pytest
from os import getenv

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.theme import ThemeBasic, ThemeKind
from ooodev.utils.info import Info


def test_theme(loader) -> None:
    if getenv("DEV_CONTAINER"):
        pytest.skip("Skip test in container: May be no theme data")
    ver = Info.version_info
    if ver < (7, 5, 0, 0):
        return

    theme = ThemeBasic(theme_name=ThemeKind.LIBRE_OFFICE)
    assert theme.comment_color >= -1
    assert theme.error_color >= -1
    assert theme.identifier_color >= -1
    assert theme.number_color >= -1
    assert theme.keyword_color >= -1
    assert theme.operator_color >= -1
    assert theme.string_color >= -1


def test_theme_default(loader) -> None:
    if getenv("DEV_CONTAINER"):
        pytest.skip("Skip test in container: May be no theme data")
    ver = Info.version_info
    if ver < (7, 5, 0, 0):
        return

    theme = ThemeBasic()
    assert theme.comment_color >= -1
    assert theme.error_color >= -1
    assert theme.identifier_color >= -1
    assert theme.number_color >= -1
    assert theme.keyword_color >= -1
    assert theme.operator_color >= -1
    assert theme.string_color >= -1


def test_theme_dark(loader) -> None:
    if getenv("DEV_CONTAINER"):
        pytest.skip("Skip test in container: May be no theme data")
    ver = Info.version_info
    if ver < (7, 5, 0, 0):
        return

    theme = ThemeBasic(theme_name=ThemeKind.LIBRE_OFFICE_DARK)
    assert theme.comment_color >= -1
    assert theme.error_color >= -1
    assert theme.identifier_color >= -1
    assert theme.number_color >= -1
    assert theme.keyword_color >= -1
    assert theme.operator_color >= -1
    assert theme.string_color >= -1
