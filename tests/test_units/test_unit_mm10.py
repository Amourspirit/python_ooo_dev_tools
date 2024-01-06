from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.units import UnitMM100
from ooodev.units import UnitInch
from ooodev.units import UnitMM10
from ooodev.units import UnitPX


def test_eq() -> None:
    u1 = UnitMM10(2.33)
    u2 = UnitMM10(2.33)
    assert u1 == u2


def test_add() -> None:
    u1 = UnitMM10(1.55)
    u2 = UnitMM10(2.24)
    u3 = u1 + u2
    assert u3.almost_equal(3.79)


def test_add_float() -> None:
    u1 = UnitMM10(1.56)
    u3 = u1 + 2.44
    assert u3.almost_equal(4)


def test_add_float_rev() -> None:
    u1 = UnitMM10(1.56)
    u3 = 2.88 + u1
    assert u3.almost_equal(4.44)


def test_add_unit_inch() -> None:
    u1 = UnitMM10(1.44)
    u2 = UnitInch(2)  # 5080 mm 100, 508 mm10
    u3 = u1 + u2
    assert u3.almost_equal(509.44)


def test_sub() -> None:
    u1 = UnitMM10(3.77)
    u2 = UnitMM10(1.29)
    u3 = u1 - u2
    assert u3.almost_equal(2.48)


def test_sub_unit_inch() -> None:
    u1 = UnitMM10(524)  # 52.4 mm
    u2 = UnitInch(2)  # 5080 mm100, 50.8 mm
    # 524 - 508 = 16
    u3 = u1 - u2
    assert u3.almost_equal(16)  # 1.6 mm


def test_sub_float() -> None:
    u1 = UnitMM10(3.5)
    u3 = u1 - 1.5
    assert u3.almost_equal(2)


def test_sub_float_rev() -> None:
    u1 = UnitMM10(3.8)
    u3 = 1.2 - u1
    assert u3.almost_equal(-2.6)


def test_mul() -> None:
    u1 = UnitMM10(3.7)
    u2 = UnitMM10(2.88)
    u3 = u1 * u2
    assert u3.almost_equal(10.656)


def test_mul_float() -> None:
    u1 = UnitMM10(3.89)
    u3 = u1 * 2.87
    assert u3.almost_equal(11.1643)


def test_mul_float_rev() -> None:
    u1 = UnitMM10(2.87)
    u3 = 3.89 * u1  # type: ignore
    assert u3.almost_equal(11.1643)  # type: ignore


def test_mul_inch() -> None:
    u1 = UnitMM10(2.54)
    u2 = UnitInch(2)  # 5080 mm 100, 50.8 mm
    u3 = u1 * u2
    assert u3.almost_equal(1290.32)


def test_mul_px() -> None:
    u1 = UnitMM10(2.54)
    u2 = UnitPX.from_inch(2)  # 5080 mm 100, 50.8 mm
    u3 = u1 * u2
    assert u3.almost_equal(1290.32)


def test_div() -> None:
    u1 = UnitMM10(6)
    u2 = UnitMM10(2)
    u3 = u1 / u2
    assert u3.almost_equal(3)


def test_div_inch() -> None:
    u1 = UnitMM10.from_inch(4)  # 5080 mm 100
    u2 = UnitInch(2)  # 5080 mm 100
    u3 = u1 / u2
    assert u3 == 2.0

    u1 = UnitMM10.from_inch(10)  # 5080 mm 100
    u2 = UnitInch(3)  # 5080 mm 100
    u3 = u1 / u2
    assert u3 == 3.3333333333333335

    u4 = UnitMM100.from_inch(2)
    u5 = u1 / u4
    assert u5 == 5.0


def test_div_float() -> None:
    u1 = UnitMM10(6)
    u3 = u1 / 2.4
    # 6 / 2.4 = 2.5
    assert u3 == 2.5


def test_div_float_rev() -> None:
    u1 = UnitMM10(2.88)
    u3 = 6 / u1
    # 6 / 2.88 = 2.0833333333333335
    assert u3 == 2.0833333333333335


def test_div_by_zero_int() -> None:
    u1 = UnitMM10(2)
    with pytest.raises(ZeroDivisionError):
        _ = u1 / 0


def test_div_by_zero_unit() -> None:
    u1 = UnitMM10(2)
    u2 = UnitMM10(0)
    with pytest.raises(ZeroDivisionError):
        _ = u1 / u2


def test_less_than() -> None:
    u1 = UnitMM10(2.9)
    u2 = UnitMM10(3)
    assert u1 < u2


def test_less_than_float() -> None:
    u1 = UnitMM10(2.99)
    assert u1 < 3.0


def test_less_than_float_rev() -> None:
    u1 = UnitMM10(3)
    assert 2.99 < u1


def test_less_than_equal() -> None:
    u1 = UnitMM10(2.8)
    u2 = UnitMM10(3.1)
    assert u1 <= u2
    u2 = UnitMM10(2.8)
    assert u1 <= u2


def test_less_than_equal_float() -> None:
    u1 = UnitMM10(2.72)
    assert u1 <= 3.0
    assert u1 <= 2.72


def test_less_than_equal_float_rev() -> None:
    u1 = UnitMM10(2.4)
    assert 1.2 <= u1
    assert 2.4 <= u1


def test_greater_than() -> None:
    u1 = UnitMM10(3.0)
    u2 = UnitMM10(2.8)
    assert u1 > u2


def test_greater_than_float() -> None:
    u1 = UnitMM10(3.1)
    assert u1 > 2.9


def test_greater_than_float_rev() -> None:
    u1 = UnitMM10(3.3)
    assert 3.4 > u1


def test_greater_than_equal() -> None:
    u1 = UnitMM10(331.2)
    u2 = UnitMM10(330.0)
    assert u1 >= u2
    u2 = UnitMM10(331.2)
    assert u1 >= u2


def test_greater_than_equal_float() -> None:
    u1 = UnitMM10(3.3)
    assert u1 >= 2.2
    assert u1 >= 3.3
    assert u1 >= 3


def test_greater_than_equal_float_rev() -> None:
    u1 = UnitMM10(3.3)
    assert 4.1 >= u1
    assert 3.3 >= u1
    assert 4 >= u1


def test_not_equal() -> None:
    u1 = UnitMM10(2.1)
    u2 = UnitMM10(2.01)
    assert u1 != u2


def test_not_equal_float() -> None:
    u1 = UnitMM10(2.1)
    assert u1 != 3.1
    assert u1 != 3


def test_not_equal_float_rev() -> None:
    u1 = UnitMM10(2)
    assert 3.1 != u1
    assert 3 != u1


def test_plus_equals() -> None:
    u1 = UnitMM10(2)
    u1 += 3
    assert isinstance(u1, UnitMM10)
    assert u1 == 5
    u2 = UnitInch.from_mm10(2)
    u1 += u2
    assert isinstance(u1, UnitMM10)
    assert u1 == 7


def test_minus_equals() -> None:
    u1 = UnitMM10(5)
    u1 -= 1
    assert isinstance(u1, UnitMM10)
    assert u1 == 4

    u2 = UnitInch.from_mm10(2)
    u1 -= u2
    assert isinstance(u1, UnitMM10)
    assert u1 == 2


def test_mul_equals() -> None:
    u1 = UnitMM10(2)
    u1 *= 3
    assert isinstance(u1, UnitMM10)
    assert u1 == 6

    u2 = UnitInch.from_mm10(2)
    u1 *= u2
    assert isinstance(u1, UnitMM10)
    assert u1 == 12


def test_div_equals() -> None:
    u1 = UnitMM10(12)
    u1 /= 2
    assert isinstance(u1, UnitMM10)
    assert u1 == 6

    u2 = UnitInch.from_mm10(2)
    u1 /= u2
    assert isinstance(u1, UnitMM10)
    assert u1 == 3
