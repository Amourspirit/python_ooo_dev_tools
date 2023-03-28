from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.theme import ThemeGeneral, ThemeKind
from ooodev.utils.info import Info


def test_theme(loader) -> None:
    ver = Info.version_info
    if ver < (7, 5, 0, 0):
        return

    theme = ThemeGeneral(theme_name=ThemeKind.LIBRE_OFFICE)
    assert theme.background_color >= -1
    assert theme.font_color >= -1
    assert theme.links_color >= -1
    assert isinstance(theme.links_visible, bool)
    assert theme.links_visited_color >= -1
    assert isinstance(theme.links_visited_visible, bool)
    assert theme.object_boundaries_color >= -1
    assert theme.shadow_color >= -1
    assert isinstance(theme.shadow_visible, bool)
    assert theme.smart_tags_color >= -1
    assert theme.table_boundaries_color >= -1
    assert isinstance(theme.table_boundaries_visible, bool)


def test_theme_default(loader) -> None:
    ver = Info.version_info
    if ver < (7, 5, 0, 0):
        return

    theme = ThemeGeneral()
    assert theme.background_color >= -1
    assert theme.font_color >= -1
    assert theme.links_color >= -1
    assert isinstance(theme.links_visible, bool)
    assert theme.links_visited_color >= -1
    assert isinstance(theme.links_visited_visible, bool)
    assert theme.object_boundaries_color >= -1
    assert theme.shadow_color >= -1
    assert isinstance(theme.shadow_visible, bool)
    assert theme.smart_tags_color >= -1
    assert theme.table_boundaries_color >= -1
    assert isinstance(theme.table_boundaries_visible, bool)


def test_theme_dark(loader) -> None:
    ver = Info.version_info
    if ver < (7, 5, 0, 0):
        return
    theme = ThemeGeneral(theme_name=ThemeKind.LIBRE_OFFICE_DARK)
    assert theme.background_color >= -1
    assert theme.font_color >= -1
    assert theme.links_color >= -1
    assert isinstance(theme.links_visible, bool)
    assert theme.links_visited_color >= -1
    assert isinstance(theme.links_visited_visible, bool)
    assert theme.object_boundaries_color >= -1
    assert theme.shadow_color >= 1
    assert isinstance(theme.shadow_visible, bool)
    assert theme.smart_tags_color >= -1
    assert theme.table_boundaries_color >= -1
    assert isinstance(theme.table_boundaries_visible, bool)
