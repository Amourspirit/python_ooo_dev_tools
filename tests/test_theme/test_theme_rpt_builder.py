from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.theme import ThemeRptBuilder
from ooodev.utils.info import Info
from ooodev.write import WriteDoc


def test_theme_default(loader) -> None:
    ver = Info.version_info
    if ver < (7, 6, 0, 0):
        return
    doc = None
    try:
        doc = WriteDoc.create_doc()
        theme = ThemeRptBuilder()
        assert theme.column_footer_color >= 0
        assert theme.column_header_color >= 0
        assert theme.detail_color >= 0
        assert theme.group_footer_color >= 0
        assert theme.group_header_color >= 0
        assert theme.overlap_control_color >= 0
        assert theme.page_footer_color >= 0
        assert theme.page_header_color >= 0
        assert theme.report_footer_color >= 0
        assert theme.report_header_color >= 0
        assert theme.text_box_bound_content_color >= 0
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
            _ = ThemeRptBuilder("some random name that does not exist")
    finally:
        if doc:
            doc.close()
