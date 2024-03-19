from __future__ import annotations
import pytest

# pylint: disable=import-outside-toplevel

if __name__ == "__main__":
    pytest.main([__file__])


@pytest.fixture(scope="module")
def fix_convert_pixel_to_app_font(loader):
    from ooodev.loader.lo import Lo

    # need to create a doc to get the current lo primed
    from ooodev.write import WriteDoc

    _ = WriteDoc.create_doc(loader=loader)
    # from ooodev.calc import CalcDoc

    # _ = CalcDoc.create_doc(loader=loader)
    from ooodev.adapter.awt.unit_conversion_comp import UnitConversionComp
    from ooo.dyn.awt.point import Point

    comp = UnitConversionComp(Lo.current_lo)

    def get_point(x: int, y: int) -> Point:
        nonlocal Point, comp
        p = Point(x, y)
        return comp.convert_point_to_logic(p, 17)

    return get_point


def _test_from_px(fix_convert_pixel_to_app_font, fix_almost_equal) -> None:
    from ooodev.loader import Lo

    p = fix_convert_pixel_to_app_font(100_000, 100_000)
    assert p.X != p.Y
    x_ratio = p.X / 100_000  # 0.51948
    y_ratio = p.Y / 100_000  # 0.61538
    lo_x_ratio = Lo.current_lo.app_font_pixel_ratio[0]
    lo_y_ratio = Lo.current_lo.app_font_pixel_ratio[1]
    assert fix_almost_equal(x_ratio, lo_x_ratio)
    assert fix_almost_equal(y_ratio, lo_y_ratio)
    assert Lo.current_lo.sys_font_pixel_ratio[0] is not None
    p = fix_convert_pixel_to_app_font(10, 10)
    assert p is not None
