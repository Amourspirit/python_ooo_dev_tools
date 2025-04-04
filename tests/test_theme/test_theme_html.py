from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.theme import ThemeHtml
from ooodev.utils.info import Info
from ooodev.write import WriteDoc


def test_theme_default(loader) -> None:
    ver = Info.version_info
    if ver < (7, 6, 0, 0):
        return

    doc = None
    try:
        doc = WriteDoc.create_doc()
        theme = ThemeHtml()
        assert theme.comment_color >= 0
        assert theme.keyword_color >= 0
        assert theme.sgml_color >= 0
        assert theme.unknown_color >= 0
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
            _ = ThemeHtml("some random name that does not exist")
    finally:
        if doc:
            doc.close()
