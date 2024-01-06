from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.units import UnitMM100
from ooodev.units import UnitInch


def test_eq() -> None:
    u1 = UnitMM100(2)
    u2 = UnitMM100(2)
    assert u1 == u2


def test_add() -> None:
    u1 = UnitMM100(1)
    u2 = UnitMM100(2)
    u3 = u1 + u2
    assert u3.value == 3


def test_add_int() -> None:
    u1 = UnitMM100(1)
    u3 = u1 + 2
    assert u3.value == 3


def test_add_int_rev() -> None:
    u1 = UnitMM100(1)
    u3 = 2 + u1
    assert u3.value == 3


def test_add_unit_inch() -> None:
    u1 = UnitMM100(1)
    u2 = UnitInch(2)  # 5080 mm 100
    u3 = u1 + u2
    assert u3.value == 5081


def test_sub() -> None:
    u1 = UnitMM100(3)
    u2 = UnitMM100(1)
    u3 = u1 - u2
    assert u3.value == 2


def test_sub_unit_inch() -> None:
    u1 = UnitMM100(5100)
    u2 = UnitInch(2)  # 5080 mm 100
    u3 = u1 - u2
    assert u3.value == 20


def test_sub_int() -> None:
    u1 = UnitMM100(3)
    u3 = u1 - 1
    assert u3.value == 2


def test_sub_int_rev() -> None:
    u1 = UnitMM100(3)
    u3 = 1 - u1
    assert u3.value == -2


def test_mul() -> None:
    u1 = UnitMM100(3)
    u2 = UnitMM100(2)
    u3 = u1 * u2
    assert u3.value == 6


def test_mul_int() -> None:
    u1 = UnitMM100(3)
    u3 = u1 * 2
    assert u3.value == 6


def test_mul_int_rev() -> None:
    u1 = UnitMM100(3)
    u3 = 2 * u1
    assert u3.value == 6


def test_mul_inch() -> None:
    u1 = UnitMM100(2)
    u2 = UnitInch(2)  # 5080 mm 100
    u3 = u1 * u2  # 10160 mm 100
    assert u3.value == 10160


def test_div_unit_mm_100() -> None:
    u1 = UnitMM100(6)
    u2 = UnitMM100(2)
    u3 = u1 / u2
    assert u3.value == 3


def test_div_inch() -> None:
    u1 = UnitMM100(10160)
    u2 = UnitInch(2)  # 5080 mm 100
    u3 = u1 / u2
    assert u3.value == 2


def test_div_int() -> None:
    u1 = UnitMM100(6)
    u3 = u1 / 2
    assert u3.value == 3


def test_div_int_rev() -> None:
    u1 = UnitMM100(2)
    u3 = 6 / u1
    assert u3.value == 3


def test_div_by_zero_int() -> None:
    u1 = UnitMM100(2)
    with pytest.raises(ZeroDivisionError):
        _ = u1 / 0


def test_div_by_zero_unit() -> None:
    u1 = UnitMM100(2)
    u2 = UnitMM100(0)
    with pytest.raises(ZeroDivisionError):
        _ = u1 / u2


def test_less_than() -> None:
    u1 = UnitMM100(2)
    u2 = UnitMM100(3)
    assert u1 < u2


def test_less_than_int() -> None:
    u1 = UnitMM100(2)
    assert u1 < 3


def test_less_than_int_rev() -> None:
    u1 = UnitMM100(3)
    assert 2 < u1


def test_less_than_equal() -> None:
    u1 = UnitMM100(2)
    u2 = UnitMM100(3)
    assert u1 <= u2
    u2 = UnitMM100(2)
    assert u1 <= u2


def test_less_than_equal_int() -> None:
    u1 = UnitMM100(2)
    assert u1 <= 3
    assert u1 <= 2


def test_less_than_equal_int_rev() -> None:
    u1 = UnitMM100(2)
    assert 1 <= u1
    assert 2 <= u1


def test_greater_than() -> None:
    u1 = UnitMM100(3)
    u2 = UnitMM100(2)
    assert u1 > u2


def test_greater_than_int() -> None:
    u1 = UnitMM100(3)
    assert u1 > 2


def test_greater_than_int_rev() -> None:
    u1 = UnitMM100(3)
    assert 4 > u1


def test_greater_than_equal() -> None:
    u1 = UnitMM100(3)
    u2 = UnitMM100(2)
    assert u1 >= u2
    u2 = UnitMM100(3)
    assert u1 >= u2


def test_greater_than_equal_int() -> None:
    u1 = UnitMM100(3)
    assert u1 >= 2
    assert u1 >= 3


def test_greater_than_equal_int_rev() -> None:
    u1 = UnitMM100(3)
    assert 4 >= u1
    assert 3 >= u1


def test_not_equal() -> None:
    u1 = UnitMM100(2)
    u2 = UnitMM100(3)
    assert u1 != u2


def test_not_equal_int() -> None:
    u1 = UnitMM100(2)
    assert u1 != 3


def test_not_equal_int_rev() -> None:
    u1 = UnitMM100(2)
    assert 3 != u1


def test_plus_equals() -> None:
    u1 = UnitMM100(2)
    u1 += 3
    assert isinstance(u1, UnitMM100)
    assert u1 == 5
    u2 = UnitInch.from_mm100(2)
    u1 += u2
    assert isinstance(u1, UnitMM100)
    assert u1 == 7


def test_minus_equals() -> None:
    u1 = UnitMM100(5)
    u1 -= 1
    assert isinstance(u1, UnitMM100)
    assert u1 == 4

    u2 = UnitInch.from_mm100(2)
    u1 -= u2
    assert isinstance(u1, UnitMM100)
    assert u1 == 2


def test_mul_equals() -> None:
    u1 = UnitMM100(2)
    u1 *= 3
    assert isinstance(u1, UnitMM100)
    assert u1 == 6

    u2 = UnitInch.from_mm100(2)
    u1 *= u2
    assert isinstance(u1, UnitMM100)
    assert u1 == 12


def test_div_equals() -> None:
    u1 = UnitMM100(12)
    u1 /= 2
    assert isinstance(u1, UnitMM100)
    assert u1 == 6

    u2 = UnitInch.from_mm100(2)
    u1 /= u2
    assert isinstance(u1, UnitMM100)
    assert u1 == 3
