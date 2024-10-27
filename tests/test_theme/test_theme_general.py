from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.theme import ThemeGeneral
from ooodev.utils.info import Info
from ooodev.write import WriteDoc


def test_theme_default(loader) -> None:
    ver = Info.version_info
    if ver < (7, 6, 0, 0):
        return
    doc = None
    try:
        doc = WriteDoc.create_doc()
        theme = ThemeGeneral()
        assert theme.background_color >= 0
        assert theme.font_color >= 0
        assert theme.links_color >= 0
        assert isinstance(theme.links_visible, bool)
        assert theme.links_visited_color >= 0
        assert isinstance(theme.links_visited_visible, bool)
        assert theme.object_boundaries_color >= 0
        assert theme.shadow_color >= 0
        assert isinstance(theme.shadow_visible, bool)
        assert theme.smart_tags_color >= 0
        assert theme.table_boundaries_color >= 0
        assert isinstance(theme.table_boundaries_visible, bool)
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
            _ = ThemeGeneral("some random name that does not exist")
    finally:
        if doc:
            doc.close()
