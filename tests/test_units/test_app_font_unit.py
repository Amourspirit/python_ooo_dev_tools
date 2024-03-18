from __future__ import annotations
import pytest

# pylint: disable=import-outside-toplevel
from unittest.mock import patch

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.units import UnitPX
from ooodev.units import UnitAppFont
from ooodev.units import UnitLength

# there is a relationship between pixels and app font size.
# The app font size ratio is app_font_pixel_ratio * pixel_size


def test_from_px(loader) -> None:
    with patch("ooodev.loader.lo.Lo._lo_inst._app_font_pixel_ratio", 0.5):
        unit = UnitAppFont.from_px(2)
        assert unit.value == 4.0

        assert unit.get_value_px() == 2.0


def test_from_unit_val(loader) -> None:
    with patch("ooodev.loader.lo.Lo._lo_inst._app_font_pixel_ratio", 0.5):
        px = UnitPX(2)
        unit = UnitAppFont.from_unit_val(px)
        assert unit.value == 4.0

        assert unit.get_value_px() == 2.0

        unit = UnitAppFont.from_unit_val(4)
        assert unit.value == 4.0

        assert unit.get_value_px() == 2.0


def test_convert_to(loader) -> None:
    with patch("ooodev.loader.lo.Lo._lo_inst._app_font_pixel_ratio", 0.5):
        unit = UnitAppFont(4.0)
        val = unit.convert_to(UnitLength.PX)
        assert val == 2.0

        assert unit.get_value_px() == 2.0


def test_eq(loader) -> None:
    with patch("ooodev.loader.lo.Lo._lo_inst._app_font_pixel_ratio", 0.5):
        unit = UnitAppFont(4.0)
        unit_px = UnitPX(2)
        assert unit == unit_px
        assert unit == 4.0
        assert unit == UnitAppFont(4.0)
        assert unit != UnitAppFont(4.2)


def test_lt(loader) -> None:
    with patch("ooodev.loader.lo.Lo._lo_inst._app_font_pixel_ratio", 0.5):
        unit = UnitAppFont(4.0)
        unit_px = UnitPX(2.1)
        assert unit < unit_px
        assert unit < 4.1
        assert unit < UnitAppFont(4.01)


def test_gt(loader) -> None:
    with patch("ooodev.loader.lo.Lo._lo_inst._app_font_pixel_ratio", 0.5):
        unit = UnitAppFont(4.0)
        unit_px = UnitPX(1.9)
        assert unit > unit_px
        assert unit > 3.9
        assert unit > UnitAppFont(3.99)


def test_lt_eq(loader) -> None:
    with patch("ooodev.loader.lo.Lo._lo_inst._app_font_pixel_ratio", 0.5):
        unit = UnitAppFont(4.0)
        unit_px = UnitPX(2.1)
        assert unit <= unit_px
        assert unit <= 4.1
        assert unit <= UnitAppFont(4.01)

        assert unit <= UnitPX(2)
        assert unit <= 4
        assert unit <= UnitAppFont(4.0)


def test_gt_eq(loader) -> None:
    with patch("ooodev.loader.lo.Lo._lo_inst._app_font_pixel_ratio", 0.5):
        unit = UnitAppFont(4.0)
        unit_px = UnitPX(1.9)
        assert unit >= unit_px
        assert unit >= 3.9
        assert unit >= UnitAppFont(3.99)

        unit_px = UnitPX(2)
        assert unit >= unit_px
        assert unit >= 4
        assert unit >= UnitAppFont(4)


def test_add(loader) -> None:
    u1 = UnitAppFont(1.55)
    u2 = UnitAppFont(2.24)
    u3 = u1 + u2
    assert u3.almost_equal(3.79)


def test_add_px(loader) -> None:
    with patch("ooodev.loader.lo.Lo._lo_inst._app_font_pixel_ratio", 0.5):
        u1 = UnitAppFont(2)
        u2 = UnitPX(2)
        u3 = u1 + u2
        assert u3.almost_equal(6)


def test_add_float(loader) -> None:
    u1 = UnitAppFont(1.56)
    u3 = u1 + 2.44
    assert u3.almost_equal(4)


def test_add_float_rev(loader) -> None:
    u1 = UnitAppFont(1.56)
    u3 = 2.88 + u1
    assert u3.almost_equal(4.44)


def test_sub(loader) -> None:
    u1 = UnitAppFont(3.77)
    u2 = UnitAppFont(1.29)
    u3 = u1 - u2
    assert u3.almost_equal(2.48)


def test_sub_float(loader) -> None:
    u1 = UnitAppFont(3.5)
    u3 = u1 - 1.5
    assert u3.almost_equal(2)


def test_sub_px(loader) -> None:
    with patch("ooodev.loader.lo.Lo._lo_inst._app_font_pixel_ratio", 0.5):
        u1 = UnitAppFont(10)
        u2 = UnitPX(4)
        u3 = u1 - u2
        assert u3.almost_equal(2)


def test_mul(loader) -> None:
    u1 = UnitAppFont(3.7)
    u2 = UnitAppFont(2.88)
    u3 = u1 * u2
    assert u3.almost_equal(10.656)


def test_mul_float(loader) -> None:
    u1 = UnitAppFont(3.89)
    u3 = u1 * 2.87
    assert u3.almost_equal(11.1643)


def test_mul_px(loader) -> None:
    with patch("ooodev.loader.lo.Lo._lo_inst._app_font_pixel_ratio", 0.5):
        u1 = UnitAppFont(10)
        u2 = UnitPX(2)
        u3 = u1 * u2
        assert u3.almost_equal(40)


def test_mul_float_rev(loader) -> None:
    u1 = UnitAppFont(2.87)
    u3 = 3.89 * u1  # type: ignore
    assert u3.almost_equal(11.1643)  # type: ignore


def test_div(loader) -> None:
    u1 = UnitAppFont(6)
    u2 = UnitAppFont(2)
    u3 = u1 / u2
    assert u3.almost_equal(3)


def test_div_px(loader) -> None:
    with patch("ooodev.loader.lo.Lo._lo_inst._app_font_pixel_ratio", 0.5):
        u1 = UnitAppFont.from_px(4)
        u2 = UnitPX(2)
        u3 = u1 / u2
        assert u3 == 2.0


def test_div_float(loader) -> None:
    u1 = UnitAppFont(6)
    u3 = u1 / 2.4
    assert u3 == 2.5


def test_div_float_rev(loader) -> None:
    u1 = UnitAppFont(2.88)
    u3 = 6 / u1
    assert u3 == 2.0833333333333335


def test_div_by_zero_int(loader) -> None:
    u1 = UnitAppFont(2)
    with pytest.raises(ZeroDivisionError):
        _ = u1 / 0


def test_div_by_zero_unit(loader) -> None:
    u1 = UnitAppFont(2)
    u2 = UnitAppFont(0)
    with pytest.raises(ZeroDivisionError):
        _ = u1 / u2
