from __future__ import annotations
import pytest
from os import getenv

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.theme import ThemeHtml, ThemeKind
from ooodev.utils.info import Info


def test_theme(loader) -> None:
    ver = Info.version_info
    if ver < (7, 5, 0, 0):
        return

    theme = ThemeHtml(theme_name=ThemeKind.LIBRE_OFFICE)
    assert theme.comment_color >= -1
    assert theme.keyword_color >= -1
    assert theme.sgml_color >= -1
    assert theme.unknown_color >= -1


def test_theme_default(loader) -> None:
    ver = Info.version_info
    if ver < (7, 5, 0, 0):
        return

    theme = ThemeHtml()
    assert theme.comment_color >= -1
    assert theme.keyword_color >= -1
    assert theme.sgml_color >= -1
    assert theme.unknown_color >= -1


def test_theme_dark(loader) -> None:
    ver = Info.version_info
    if ver < (7, 5, 0, 0):
        return

    theme = ThemeHtml(theme_name=ThemeKind.LIBRE_OFFICE_DARK)
    assert theme.comment_color >= -1
    assert theme.keyword_color >= -1
    assert theme.sgml_color >= -1
    assert theme.unknown_color >= -1
