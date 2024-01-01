from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.units import UnitMM100
from ooodev.units import UnitInch
from ooodev.units import UnitInch1000


def test_eq() -> None:
    u1 = UnitInch1000(2)
    u2 = UnitInch1000(2)
    assert u1 == u2


def test_add() -> None:
    u1 = UnitInch1000(1)
    u2 = UnitInch1000(2)
    u3 = u1 + u2
    assert u3.value == 3


def test_add_int() -> None:
    u1 = UnitInch1000(1)
    u3 = u1 + 2
    assert u3.value == 3


def test_add_int_rev() -> None:
    u1 = UnitInch1000(1)
    u3 = 2 + u1
    assert u3.value == 3


def test_add_unit_inch() -> None:
    u1 = UnitInch1000(1)
    u2 = UnitInch(2)  # 5080 mm 100
    u3 = u1 + u2
    assert u3.value == 2001


def test_sub() -> None:
    u1 = UnitInch1000(3)
    u2 = UnitInch1000(1)
    u3 = u1 - u2
    assert u3.value == 2


def test_sub_unit_inch() -> None:
    u1 = UnitInch1000(2001)
    u2 = UnitInch(2)  # 5080 mm 100
    u3 = u1 - u2
    assert u3.value == 1

    u1 = UnitInch1000(1)
    u3 = u1 - u2
    assert u3.value == -1999


def test_sub_int() -> None:
    u1 = UnitInch1000(3)
    u3 = u1 - 1
    assert u3.value == 2


def test_sub_int_rev() -> None:
    u1 = UnitInch1000(3)
    u3 = 1 - u1
    assert u3.value == -2


def test_mul() -> None:
    u1 = UnitInch1000(3)
    u2 = UnitInch1000(2)
    u3 = u1 * u2
    assert u3.value == 6


def test_mul_int() -> None:
    u1 = UnitInch1000(3)
    u3 = u1 * 2
    assert u3.value == 6


def test_mul_int_rev() -> None:
    u1 = UnitInch1000(3)
    u3 = 2 * u1
    assert u3.value == 6


def test_mul_inch() -> None:
    u1 = UnitInch1000(2)  # 5 mm 100
    u2 = UnitInch(2)  # 5080 mm 100
    u3 = u1 * u2
    assert u3.value == 4000


def test_div() -> None:
    u1 = UnitInch1000(6)
    u2 = UnitInch1000(2)
    u3 = u1 / u2
    assert u3.value == 3


def test_div_inch() -> None:
    u1 = UnitInch1000.from_inch(4)  # 5080 mm 100
    u2 = UnitInch(2)  # 5080 mm 100
    u3 = u1 / u2
    assert u3.value == 2

    u4 = UnitMM100.from_inch(2)
    u5 = u1 / u4
    assert u5.value == 2


def test_div_int() -> None:
    u1 = UnitInch1000(6)
    u3 = u1 / 2
    assert u3.value == 3


def test_div_int_rev() -> None:
    u1 = UnitInch1000(2)
    u3 = 6 / u1
    assert u3.value == 3


def test_div_by_zero_int() -> None:
    u1 = UnitInch1000(2)
    with pytest.raises(ZeroDivisionError):
        _ = u1 / 0


def test_div_by_zero_unit() -> None:
    u1 = UnitInch1000(2)
    u2 = UnitInch1000(0)
    with pytest.raises(ZeroDivisionError):
        _ = u1 / u2


def test_less_than() -> None:
    u1 = UnitInch1000(2)
    u2 = UnitInch1000(3)
    assert u1 < u2


def test_less_than_int() -> None:
    u1 = UnitInch1000(2)
    assert u1 < 3


def test_less_than_int_rev() -> None:
    u1 = UnitInch1000(3)
    assert 2 < u1


def test_less_than_equal() -> None:
    u1 = UnitInch1000(2)
    u2 = UnitInch1000(3)
    assert u1 <= u2
    u2 = UnitInch1000(2)
    assert u1 <= u2


def test_less_than_equal_int() -> None:
    u1 = UnitInch1000(2)
    assert u1 <= 3
    assert u1 <= 2


def test_less_than_equal_int_rev() -> None:
    u1 = UnitInch1000(2)
    assert 1 <= u1
    assert 2 <= u1


def test_greater_than() -> None:
    u1 = UnitInch1000(3)
    u2 = UnitInch1000(2)
    assert u1 > u2


def test_greater_than_int() -> None:
    u1 = UnitInch1000(3)
    assert u1 > 2


def test_greater_than_int_rev() -> None:
    u1 = UnitInch1000(3)
    assert 4 > u1


def test_greater_than_equal() -> None:
    u1 = UnitInch1000(3)
    u2 = UnitInch1000(2)
    assert u1 >= u2
    u2 = UnitInch1000(3)
    assert u1 >= u2


def test_greater_than_equal_int() -> None:
    u1 = UnitInch1000(3)
    assert u1 >= 2
    assert u1 >= 3


def test_greater_than_equal_int_rev() -> None:
    u1 = UnitInch1000(3)
    assert 4 >= u1
    assert 3 >= u1


def test_not_equal() -> None:
    u1 = UnitInch1000(2)
    u2 = UnitInch1000(3)
    assert u1 != u2


def test_not_equal_int() -> None:
    u1 = UnitInch1000(2)
    assert u1 != 3


def test_not_equal_int_rev() -> None:
    u1 = UnitInch1000(2)
    assert 3 != u1


def test_plus_equals() -> None:
    u1 = UnitInch1000(2)
    u1 += 3
    assert isinstance(u1, UnitInch1000)
    assert u1 == 5
    u2 = UnitInch.from_inch1000(2)
    u1 += u2
    assert isinstance(u1, UnitInch1000)
    assert u1 == 7


def test_minus_equals() -> None:
    u1 = UnitInch1000(5)
    u1 -= 1
    assert isinstance(u1, UnitInch1000)
    assert u1 == 4

    u2 = UnitInch.from_inch1000(2)
    u1 -= u2
    assert isinstance(u1, UnitInch1000)
    assert u1 == 2


def test_mul_equals() -> None:
    u1 = UnitInch1000(2)
    u1 *= 3
    assert isinstance(u1, UnitInch1000)
    assert u1 == 6

    u2 = UnitInch.from_inch1000(2)
    u1 *= u2
    assert isinstance(u1, UnitInch1000)
    assert u1 == 12


def test_div_equals() -> None:
    u1 = UnitInch1000(12)
    u1 /= 2
    assert isinstance(u1, UnitInch1000)
    assert u1 == 6

    u2 = UnitInch.from_inch1000(2)
    u1 /= u2
    assert isinstance(u1, UnitInch1000)
    assert u1 == 3
