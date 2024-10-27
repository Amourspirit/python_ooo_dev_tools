from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.theme import ThemeDraw, ThemeKind
from ooodev.utils.info import Info
from ooodev.write import WriteDoc


def test_theme_default(loader) -> None:
    ver = Info.version_info
    if ver < (7, 6, 0, 0):
        return

    theme = ThemeDraw()
    assert theme.grid_color >= 0
    assert isinstance(theme.grid_visible, bool)


def test_theme_no_exist(loader) -> None:
    ver = Info.version_info
    if ver < (7, 6, 0, 0):
        return

    doc = None
    try:
        doc = WriteDoc.create_doc()
        with pytest.raises(ValueError):
            _ = ThemeDraw("some random name that does not exist")
    finally:
        if doc:
            doc.close()
