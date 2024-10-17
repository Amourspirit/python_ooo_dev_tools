from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.format.writer.modify.frame.wrap import (
    Options,
    InnerOptions,
    StyleFrameKind,
    Settings,
    InnerSettings,
    WrapTextMode,
)
from ooodev.format import Styler
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.write import Write


def test_write(loader, para_text) -> None:
    # this test may fail if test is run in debug mode.
    # this seems to be VS Code thing. Whe an objects attribute is deleted
    # eg: style.prop_inner = do
    # In Debug mode VS Code is holding on to the deleted attribue value for some reason.
    # style.prop_inner = do
    # >           assert style.prop_inner is do
    # E           assert <ooodev.format.direct.frame.wrap.options.Options object ...
    # Test passes fine in regualr mode

    # delay = 0 if Lo.bridge_connector.headless else 3_000
    delay = 0

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)
    try:
        cursor = Write.get_cursor(doc)
        if not Lo.bridge_connector.headless:
            Write.append_para(cursor=cursor, text=para_text)

        style_settings = Settings(mode=WrapTextMode.THROUGH, style_name=StyleFrameKind.FRAME)

        # background can only be set to true when mode is Through
        style = Options(first=True, background=True, style_name=style_settings.prop_style_name)

        Styler.apply(doc, style_settings, style)
        # props = style.get_style_props(doc)

        f_style = Options.from_style(doc=doc, style_name=style.prop_style_name)
        assert f_style.prop_inner.prop_first
        assert f_style.prop_inner.prop_first == style.prop_inner.prop_first
        assert f_style.prop_inner.prop_background
        assert f_style.prop_inner.prop_background == style.prop_inner.prop_background

        ds = InnerSettings(mode=WrapTextMode.DYNAMIC)
        style_settings.prop_inner = ds

        do = InnerOptions(background=False)
        do = do.contour.outside.overlap.first
        style.prop_inner = do
        assert style.prop_inner is do
        Styler.apply(doc, style_settings, style)

        f_style = Options.from_style(doc=doc, style_name=style.prop_style_name)
        assert f_style.prop_inner.prop_first
        assert f_style.prop_inner.prop_first == style.prop_inner.prop_first
        assert f_style.prop_inner.prop_contour
        assert f_style.prop_inner.prop_contour == style.prop_inner.prop_contour
        assert f_style.prop_inner.prop_overlap
        assert f_style.prop_inner.prop_overlap == style.prop_inner.prop_overlap
        assert f_style.prop_inner.prop_outside
        assert f_style.prop_inner.prop_outside == style.prop_inner.prop_outside
        assert f_style.prop_inner.prop_background == False
        assert f_style.prop_inner.prop_background == style.prop_inner.prop_background

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
