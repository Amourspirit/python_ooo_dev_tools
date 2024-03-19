from __future__ import annotations
import pytest

# pylint: disable=import-outside-toplevel
# pylint: disable=unused-argument
from unittest.mock import patch

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.units import UnitPX
from ooodev.units import UnitMM
from ooodev.units import UnitAppFontX
from ooodev.units import UnitAppFontY
from ooodev.units import UnitLength

# there is a relationship between pixels and app font size.
# The app font size ratio is app_font_pixel_ratio * pixel_size


def test_from_px(loader) -> None:
    with patch("ooodev.loader.lo.Lo._lo_inst._app_font_pixel_ratio", (0.6, 0.5)):
        unit = UnitAppFontY.from_px(4)
        assert unit.value == 2.0

        assert unit.get_value_px() == 4.0


def test_from_unit_val(loader) -> None:
    with patch("ooodev.loader.lo.Lo._lo_inst._app_font_pixel_ratio", (0.6, 0.5)):
        px = UnitPX(10)
        unit = UnitAppFontY.from_unit_val(px)
        assert unit.value == 5.0

        assert unit.get_value_px() == 10.0

        unit = UnitAppFontY.from_unit_val(4)
        assert unit.value == 4.0

        assert unit.get_value_px() == 8.0


def test_convert_to(loader) -> None:
    with patch("ooodev.loader.lo.Lo._lo_inst._app_font_pixel_ratio", (0.6, 0.5)):
        unit = UnitAppFontY(4.0)
        val = unit.convert_to(UnitLength.PX)
        assert val == 8.0

        assert unit.get_value_px() == 8.0


def test_eq(loader) -> None:
    with patch("ooodev.loader.lo.Lo._lo_inst._app_font_pixel_ratio", (0.6, 0.5)):
        unit = UnitAppFontY(4.0)
        unit_px = UnitPX(8)
        assert unit == unit_px
        assert unit == 4.0
        assert unit == UnitAppFontY(4.0)
        assert unit != UnitAppFontY(4.2)
        assert unit != UnitAppFontX(unit.value)

        unit_mm = UnitMM.from_px(unit_px.value)
        assert unit == unit_mm


def test_lt(loader) -> None:
    with patch("ooodev.loader.lo.Lo._lo_inst._app_font_pixel_ratio", (0.6, 0.5)):
        unit = UnitAppFontY(4.0)
        unit_px = UnitPX(8.1)
        assert unit < unit_px
        assert unit < 4.1
        assert unit < UnitAppFontY(4.01)

        unit_mm = UnitMM.from_px(unit_px.value)
        assert unit < unit_mm


def test_gt(loader) -> None:
    with patch("ooodev.loader.lo.Lo._lo_inst._app_font_pixel_ratio", (0.6, 0.5)):
        unit = UnitAppFontY(4.0)
        unit_px = UnitPX(7.9)
        assert unit > unit_px
        assert unit > 3.9
        assert unit > UnitAppFontY(3.99)

        unit_mm = UnitMM.from_px(unit_px.value)
        assert unit > unit_mm


def test_lt_eq(loader) -> None:
    with patch("ooodev.loader.lo.Lo._lo_inst._app_font_pixel_ratio", (0.6, 0.5)):
        unit = UnitAppFontY(4.0)  # 8 px
        unit_px = UnitPX(8.05)
        assert unit <= unit_px
        assert unit <= 4.1
        assert unit <= UnitAppFontY(4.01)

        assert unit <= UnitPX(8.0)
        assert unit <= 4
        assert unit <= UnitAppFontY(4.0)

        unit_mm = UnitMM.from_px(unit_px.value)
        assert unit <= unit_mm


def test_gt_eq(loader) -> None:
    with patch("ooodev.loader.lo.Lo._lo_inst._app_font_pixel_ratio", (0.6, 0.5)):
        unit = UnitAppFontY(4.0)
        unit_px = UnitPX(1.9)
        assert unit >= unit_px
        assert unit >= 3.9
        assert unit >= UnitAppFontY(3.99)

        unit_px = UnitPX(2)
        assert unit >= unit_px
        assert unit >= 4
        assert unit >= UnitAppFontY(4)

        unit_mm = UnitMM.from_px(unit_px.value)
        assert unit >= unit_mm


def test_add(loader) -> None:
    u1 = UnitAppFontY(1.55)
    u2 = UnitAppFontY(2.24)
    u3 = u1 + u2
    assert u3.almost_equal(3.79)


def test_add_px(loader) -> None:
    with patch("ooodev.loader.lo.Lo._lo_inst._app_font_pixel_ratio", (0.6, 0.5)):
        u1 = UnitAppFontY(2)  # 4 px
        u2 = UnitPX(2)
        u3 = u1 + u2  # add 2 px, that is 1 app font units
        unit_px = UnitPX(u3.get_value_px())
        assert unit_px.almost_equal(6)
        assert u3.almost_equal(3)

        unit_mm = UnitMM.from_px(u2.value)
        u3 = u1 + unit_mm
        assert u3 == 3.0


def test_add_float(loader) -> None:
    u1 = UnitAppFontY(1.56)
    u3 = u1 + 2.44
    assert u3.almost_equal(4)


def test_add_float_rev(loader) -> None:
    u1 = UnitAppFontY(1.56)
    u3 = 2.88 + u1
    assert u3.almost_equal(4.44)


def test_sub(loader) -> None:
    u1 = UnitAppFontY(3.77)
    u2 = UnitAppFontY(1.29)
    u3 = u1 - u2
    assert u3.almost_equal(2.48)


def test_sub_float(loader) -> None:
    u1 = UnitAppFontY(3.5)
    u3 = u1 - 1.5
    assert u3.almost_equal(2)


def test_sub_px(loader) -> None:
    with patch("ooodev.loader.lo.Lo._lo_inst._app_font_pixel_ratio", (0.6, 0.5)):
        u1 = UnitAppFontY(10)
        u2 = UnitPX(4)  # 2 app font units
        u3 = u1 - u2
        assert u3.almost_equal(8)

        unit_mm = UnitMM.from_px(u2.value)
        u3 = u1 - unit_mm
        assert u3 == 8


def test_mul(loader) -> None:
    u1 = UnitAppFontY(3.7)
    u2 = UnitAppFontY(2.88)
    u3 = u1 * u2
    assert u3.almost_equal(10.656)


def test_mul_float(loader) -> None:
    u1 = UnitAppFontY(3.89)
    u3 = u1 * 2.87
    assert u3.almost_equal(11.1643)


def test_mul_px(loader) -> None:
    with patch("ooodev.loader.lo.Lo._lo_inst._app_font_pixel_ratio", (0.6, 0.5)):
        u1 = UnitAppFontY(4)  # 8 px, 4AF
        u2 = UnitPX(2)  # 2px,  4 AF
        unit_px = u2 * u1  # 8px * 2px = 16px
        assert unit_px.almost_equal(16)
        u3 = u1 * u2  # 8px * 2px = 16px, 4AF * 2AF = 8AF
        # 10 * 2 = 20 AP
        assert u3.almost_equal(8)  # 32 px

        u1 = UnitAppFontY(4)  # 8px, 4AF
        u2 = UnitPX(3)  # 3px,  1.5AF
        unit_px = u2 * u1  # 8px * 3px = 24px
        assert unit_px.almost_equal(24)
        u3 = u1 * u2  # 8px * 3px = 24px, 4AF * 1.5AF = 6AF
        assert u3.almost_equal(12)  # 12 px

        u1 = UnitAppFontY(400)
        u2 = UnitPX(338)
        unit_px = u2 * u1
        assert unit_px.almost_equal(270_400)
        u3 = u1 * u2
        assert u3.almost_equal(135_200)

        unit_mm = UnitMM.from_px(3)
        u1 = UnitAppFontY(4)
        u3 = u1 * unit_mm
        assert u3 == 12.0


def test_mul_float_rev(loader) -> None:
    u1 = UnitAppFontY(2.87)
    u3 = 3.89 * u1  # type: ignore
    assert u3.almost_equal(11.1643)  # type: ignore


def test_div(loader) -> None:
    u1 = UnitAppFontY(6)
    u2 = UnitAppFontY(2)
    u3 = u1 / u2
    assert u3.almost_equal(3)


def test_div_px(loader) -> None:
    with patch("ooodev.loader.lo.Lo._lo_inst._app_font_pixel_ratio", (0.6, 0.5)):
        u1 = UnitAppFontY.from_px(4)
        u2 = UnitPX(2)
        u3 = u1 / u2
        assert u3 == 2.0


def test_div_float(loader) -> None:
    u1 = UnitAppFontY(6)
    u3 = u1 / 2.4
    assert u3 == 2.5


def test_div_float_rev(loader) -> None:
    u1 = UnitAppFontY(2.88)
    u3 = 6 / u1
    assert u3 == 2.0833333333333335


def test_div_by_zero_int(loader) -> None:
    u1 = UnitAppFontY(2)
    with pytest.raises(ZeroDivisionError):
        _ = u1 / 0


def test_div_by_zero_unit(loader) -> None:
    u1 = UnitAppFontY(2)
    u2 = UnitAppFontY(0)
    with pytest.raises(ZeroDivisionError):
        _ = u1 / u2
