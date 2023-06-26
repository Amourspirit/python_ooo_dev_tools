from __future__ import annotations
import pytest
from os import getenv

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.theme import ThemeRptBuilder, ThemeKind
from ooodev.utils.info import Info


def test_theme(loader) -> None:
    if getenv("DEV_CONTAINER"):
        pytest.skip("Skip test in container: May be no theme data")
    ver = Info.version_info
    if ver < (7, 5, 0, 0):
        return

    theme = ThemeRptBuilder(theme_name=ThemeKind.LIBRE_OFFICE)
    assert theme.column_footer_color >= -1
    assert theme.column_header_color >= -1
    assert theme.detail_color >= -1
    assert theme.group_footer_color >= -1
    assert theme.group_header_color >= -1
    assert theme.overlap_control_color >= -1
    assert theme.page_footer_color >= -1
    assert theme.page_header_color >= -1
    assert theme.report_footer_color >= -1
    assert theme.report_header_color >= -1
    assert theme.text_box_bound_content_color >= -1


def test_theme_default(loader) -> None:
    if getenv("DEV_CONTAINER"):
        pytest.skip("Skip test in container: May be no theme data")
    ver = Info.version_info
    if ver < (7, 5, 0, 0):
        return

    theme = ThemeRptBuilder()
    assert theme.column_footer_color >= -1
    assert theme.column_header_color >= -1
    assert theme.detail_color >= -1
    assert theme.group_footer_color >= -1
    assert theme.group_header_color >= -1
    assert theme.overlap_control_color >= -1
    assert theme.page_footer_color >= -1
    assert theme.page_header_color >= -1
    assert theme.report_footer_color >= -1
    assert theme.report_header_color >= -1
    assert theme.text_box_bound_content_color >= -1


def test_theme_dark(loader) -> None:
    if getenv("DEV_CONTAINER"):
        pytest.skip("Skip test in container: May be no theme data")
    ver = Info.version_info
    if ver < (7, 5, 0, 0):
        return

    theme = ThemeRptBuilder(theme_name=ThemeKind.LIBRE_OFFICE_DARK)
    assert theme.column_footer_color >= -1
    assert theme.column_header_color >= -1
    assert theme.detail_color >= -1
    assert theme.group_footer_color >= -1
    assert theme.group_header_color >= -1
    assert theme.overlap_control_color >= -1
    assert theme.page_footer_color >= -1
    assert theme.page_header_color >= -1
    assert theme.report_footer_color >= -1
    assert theme.report_header_color >= -1
    assert theme.text_box_bound_content_color >= -1
