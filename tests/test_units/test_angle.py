from __future__ import annotations
from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.units.angle import Angle
from ooodev.units.angle10 import Angle10
from ooodev.units.angle100 import Angle100


# angle implements BaseIntValue so this test all dunder methods.


def test_angle() -> None:
    for i in range(361):
        angle = Angle(i)
        assert isinstance(angle, Angle)


@pytest.mark.parametrize("val", ["1", 1.2, -3.3])
def test_angle_error_incorrect_type(val: int) -> None:
    with pytest.raises(TypeError):
        Angle(val)


@pytest.mark.parametrize("start,val", [(0, 5), (1, 3), (33, 55), (255, 100)])
def test_addition_int(start: int, val: int) -> None:
    a1 = Angle(start)
    a2 = a1 + val
    assert a2.value == start + val


@pytest.mark.parametrize("start,val", [(0, 5), (1, 3), (33, 55), (255, 100)])
def test_10addition_int(start: int, val: int) -> None:
    a1 = Angle10(start)
    a2 = a1 + val
    assert a2.value == start + val


@pytest.mark.parametrize("start,val", [(0, 5), (1, 3), (33, 55), (255, 100)])
def test_100addition_int(start: int, val: int) -> None:
    a1 = Angle100(start)
    a2 = a1 + val
    assert a2.value == start + val


@pytest.mark.parametrize("start,val", [(0, 0), (0, 5), (1, 3), (33, 55), (255, 100)])
def test_addition_angle_to_angle(start: int, val: int) -> None:
    a1 = Angle(start)
    a2 = Angle(val)
    a3 = a1 + a2
    assert a3.value == start + val


@pytest.mark.parametrize("start,val", [(0, 0), (0, 5), (1, 3), (33, 55), (255, 100)])
def test_10addition_angle_to_angle(start: int, val: int) -> None:
    a1 = Angle10(start)
    a2 = Angle10(val)
    a3 = a1 + a2
    assert a3.value == start + val


@pytest.mark.parametrize("start,val", [(0, 0), (0, 5), (1, 3), (33, 55), (255, 100)])
def test_100addition_angle_to_angle(start: int, val: int) -> None:
    a1 = Angle100(start)
    a2 = Angle100(val)
    a3 = a1 + a2
    assert a3.value == start + val


@pytest.mark.parametrize("start,val", [(0, 0), (0, 5), (1, 3), (33, 55), (255, 100)])
def test_addition_plus_equal(start: int, val: int) -> None:
    a1 = Angle(start)
    a1 += val
    result = start + val
    assert a1.value == result


@pytest.mark.parametrize("start,val", [(0, 0), (0, 5), (1, 3), (33, 55), (255, 100)])
def test_10addition_plus_equal(start: int, val: int) -> None:
    a1 = Angle10(start)
    a1 += val
    result = start + val
    assert a1.value == result


@pytest.mark.parametrize("start,val", [(0, 0), (0, 500), (1000, 3000), (33, 55), (255, 100)])
def test_100addition_plus_equal(start: int, val: int) -> None:
    a1 = Angle100(start)
    a1 += val
    result = start + val
    assert a1.value == result


def test_angle_init_angle() -> None:
    a1 = Angle(2)
    a2 = Angle(int(a1))
    assert a2 == a1


def test_10angle_init_angle() -> None:
    a1 = Angle10(1000)
    a2 = Angle10(int(a1))
    assert a2 == a1


def test_100angle_init_angle() -> None:
    a1 = Angle100(10000)
    a2 = Angle100(int(a1))
    assert a2 == a1


@pytest.mark.parametrize("start,val", [(0, 0), (5, 0), (3, 1), (55, 22), (255, 100)])
def test_sub_angle_to_angle(start: int, val: int) -> None:
    a1 = Angle(start)
    a2 = Angle(val)
    a3 = a1 - a2
    assert a3.value == start - val


@pytest.mark.parametrize("start,val", [(0, 0), (5, 0), (3, 1), (55, 22), (255, 100)])
def test_10sub_angle_to_angle(start: int, val: int) -> None:
    a1 = Angle10(start)
    a2 = Angle10(val)
    a3 = a1 - a2
    assert a3.value == start - val


@pytest.mark.parametrize("start,val", [(0, 0), (5, 0), (3, 1), (55, 22), (255, 100)])
def test_100sub_angle_to_angle(start: int, val: int) -> None:
    a1 = Angle100(start)
    a2 = Angle100(val)
    a3 = a1 - a2
    assert a3.value == start - val


@pytest.mark.parametrize("start,val", [(0, 0), (5, 0), (3, 1), (55, 22), (255, 100)])
def test_minus_equal(start: int, val: int) -> None:
    a1 = Angle(start)
    a1 -= val
    result = start - val
    assert a1.value == result


@pytest.mark.parametrize("start,val", [(0, 0), (5, 0), (3, 1), (55, 22), (255, 100)])
def test_10minus_equal(start: int, val: int) -> None:
    a1 = Angle10(start)
    a1 -= val
    result = start - val
    assert a1.value == result


@pytest.mark.parametrize("start,val", [(0, 0), (5, 0), (3, 1), (55, 22), (255, 100)])
def test_100minus_equal(start: int, val: int) -> None:
    a1 = Angle100(start)
    a1 -= val
    result = start - val
    assert a1.value == result


def test_sum_angle() -> None:
    a1 = Angle(5)
    a2 = Angle(10)
    a3 = Angle(15)
    a4 = Angle(20)
    sums = sum([a1, a2, a3, a4])
    assert isinstance(sums, Angle)
    assert sums.value == 50


def test_10sum_angle() -> None:
    a1 = Angle10(5)
    a2 = Angle10(10)
    a3 = Angle10(15)
    a4 = Angle10(20)
    sums = sum([a1, a2, a3, a4])
    assert isinstance(sums, Angle10)
    assert sums.value == 50


def test_100sum_angle() -> None:
    a1 = Angle100(5)
    a2 = Angle100(10)
    a3 = Angle100(15)
    a4 = Angle100(20)
    sums = sum([a1, a2, a3, a4])
    assert isinstance(sums, Angle100)
    assert sums.value == 50


@pytest.mark.parametrize("start,val", [(0, 0), (5, 0), (3, 3), (55, 3)])
def test_mul_angle_to_angle(start: int, val: int) -> None:
    a1 = Angle(start)
    a2 = Angle(val)
    a3 = a1 * a2
    result = start * val
    assert a3.value == result


@pytest.mark.parametrize("start,val", [(0, 0), (5, 0), (3, 3), (55, 3)])
def test_10mul_angle_to_angle(start: int, val: int) -> None:
    a1 = Angle10(start)
    a2 = Angle10(val)
    a3 = a1 * a2
    result = start * val
    assert a3.value == result


@pytest.mark.parametrize("start,val", [(0, 0), (5, 0), (3, 3), (55, 3)])
def test_100mul_angle_to_angle(start: int, val: int) -> None:
    a1 = Angle100(start)
    a2 = Angle100(val)
    a3 = a1 * a2
    result = start * val
    assert a3.value == result


@pytest.mark.parametrize("start,val", [(5, 2), (300, 45), (355, 3)])
def test_div_angle_to_angle(start: int, val: int) -> None:
    a1 = Angle(start)
    a2 = Angle(val)
    a3 = a1 / a2
    result = start // val
    assert a3.value == result


@pytest.mark.parametrize("start,val", [(5, 2), (300, 45), (355, 3)])
def test_10div_angle_to_angle(start: int, val: int) -> None:
    a1 = Angle10(start)
    a2 = Angle10(val)
    a3 = a1 / a2
    result = start // val
    assert a3.value == result


@pytest.mark.parametrize("start,val", [(5, 2), (300, 45), (355, 3)])
def test_100div_angle_to_angle(start: int, val: int) -> None:
    a1 = Angle100(start)
    a2 = Angle100(val)
    a3 = a1 / a2
    result = start // val
    assert a3.value == result


@pytest.mark.parametrize("start,val", [(0, 1), (2, 3), (3, 77), (55, 198), (100, 359)])
def test_lt_angle(start: int, val: int) -> None:
    a1 = Angle(start)
    a2 = Angle(val)
    assert a1 < a2
    gt = a1 > a2
    assert gt is False


@pytest.mark.parametrize("start,val", [(0, 1), (2, 3), (3, 77), (55, 198), (100, 359)])
def test_10lt_angle(start: int, val: int) -> None:
    a1 = Angle10(start)
    a2 = Angle10(val)
    assert a1 < a2
    gt = a1 > a2
    assert gt is False


@pytest.mark.parametrize("start,val", [(0, 1), (2, 3), (3, 77), (55, 198), (100, 359)])
def test_100lt_angle(start: int, val: int) -> None:
    a1 = Angle100(start)
    a2 = Angle100(val)
    assert a1 < a2
    gt = a1 > a2
    assert gt is False


@pytest.mark.parametrize("start,val", [(0, 1), (2, 3), (5, 5), (55, 198), (360, 360)])
def test_le_angle(start: int, val: int) -> None:
    a1 = Angle(start)
    a2 = Angle(val)
    assert a1 <= a2
    gt = a1 > a2
    assert gt is False


@pytest.mark.parametrize("start,val", [(0, 1), (2, 3), (5, 5), (55, 198), (360, 360)])
def test_10le_angle(start: int, val: int) -> None:
    a1 = Angle10(start)
    a2 = Angle10(val)
    assert a1 <= a2
    gt = a1 > a2
    assert gt is False


@pytest.mark.parametrize("start,val", [(0, 1), (2, 3), (5, 5), (55, 198), (360, 360)])
def test_100le_angle(start: int, val: int) -> None:
    a1 = Angle100(start)
    a2 = Angle100(val)
    assert a1 <= a2
    gt = a1 > a2
    assert gt is False


@pytest.mark.parametrize("start,val", [(1, 0), (3, 2), (77, 3), (198, 10), (359, 3)])
def test_gt_angle(start: int, val: int) -> None:
    a1 = Angle(start)
    a2 = Angle(val)
    assert a1 > a2
    lt = a1 < a2
    assert lt is False


@pytest.mark.parametrize("start,val", [(1, 0), (3, 2), (77, 3), (198, 10), (359, 3)])
def test_10gt_angle(start: int, val: int) -> None:
    a1 = Angle10(start)
    a2 = Angle10(val)
    assert a1 > a2
    lt = a1 < a2
    assert lt is False


@pytest.mark.parametrize("start,val", [(1, 0), (3, 2), (77, 3), (198, 10), (359, 3)])
def test_100gt_angle(start: int, val: int) -> None:
    a1 = Angle100(start)
    a2 = Angle100(val)
    assert a1 > a2
    lt = a1 < a2
    assert lt is False


@pytest.mark.parametrize("start,val", [(1, 0), (3, 2), (5, 5), (198, 44), (360, 360)])
def test_ge_angle(start: int, val: int) -> None:
    a1 = Angle(start)
    a2 = Angle(val)
    assert a1 >= a2
    lt = a1 < a2
    assert lt is False


@pytest.mark.parametrize("start,val", [(1, 0), (3, 2), (5, 5), (198, 44), (360, 360)])
def test_10ge_angle(start: int, val: int) -> None:
    a1 = Angle10(start)
    a2 = Angle10(val)
    assert a1 >= a2
    lt = a1 < a2
    assert lt is False


@pytest.mark.parametrize("start,val", [(1, 0), (3, 2), (5, 5), (198, 44), (360, 360)])
def test_100ge_angle(start: int, val: int) -> None:
    a1 = Angle100(start)
    a2 = Angle100(val)
    assert a1 >= a2
    lt = a1 < a2
    assert lt is False


@pytest.mark.parametrize(("val", "expected"), [(360, 0), (-3, 357), (-181, 179), (2500, 340), (-2500, 20)])
def test_angle_weird(val: int, expected: int) -> None:
    a1 = Angle(val)
    assert a1.value == expected


@pytest.mark.parametrize(("val", "subval", "expected"), [(1, 100, 99), (2, 1, 359), (200, 201, 1)])
def test_angle_rsub(val: int, subval: int, expected: int) -> None:
    # test __rsub__
    # Test when a number subtracts an Angle
    a1 = Angle(val)
    result = subval - a1
    assert result == Angle(expected)


@pytest.mark.parametrize(("val", "subval", "expected"), [(1, 100, 99), (2, 1, 3599), (200, 201, 1)])
def test_10angle_rsub(val: int, subval: int, expected: int) -> None:
    # test __rsub__
    # Test when a number subtracts an Angle
    a1 = Angle10(val)
    result = subval - a1
    assert result == Angle10(expected)


@pytest.mark.parametrize(("val", "subval", "expected"), [(1, 100, 99), (2, 1, 35999), (200, 201, 1)])
def test_100angle_rsub(val: int, subval: int, expected: int) -> None:
    # test __rsub__
    # Test when a number subtracts an Angle
    a1 = Angle100(val)
    result = subval - a1
    assert result == Angle100(expected)


@pytest.mark.parametrize("start", [1, 3, 5, 100, 0])
def test_eq_angle(start: int) -> None:
    a1 = Angle(start)
    a2 = Angle(start)
    assert a1 == a2
    assert a2 == a1
    assert a1 == start
    assert start == a1


@pytest.mark.parametrize("start", [1, 3, 5, 100, 0])
def test_10eq_angle(start: int) -> None:
    a1 = Angle10(start)
    a2 = Angle10(start)
    assert a1 == a2
    assert a2 == a1
    assert a1 == start
    assert start == a1


@pytest.mark.parametrize("start", [1, 3, 5, 100, 0])
def test_100eq_angle(start: int) -> None:
    a1 = Angle100(start)
    a2 = Angle100(start)
    assert a1 == a2
    assert a2 == a1
    assert a1 == start
    assert start == a1


# region Angle + Angle10 + Angle100
def test_angle_plus_angle10() -> None:
    a1 = Angle(5)
    a2 = Angle10(50)
    a3 = a1 + a2
    assert isinstance(a3, Angle)
    assert a3 == 10

    a1 = Angle(90)
    a2 = Angle10(110)
    a3 = a1 + a2
    assert isinstance(a3, Angle)
    assert a3 == 101


def test_angle_plus_angle100() -> None:
    a1 = Angle(5)
    a2 = Angle100(500)
    a3 = a1 + a2
    assert isinstance(a3, Angle)
    assert a3 == 10


def test_angle10_plus_angle() -> None:
    a1 = Angle10(50)
    a2 = Angle(5)
    a3 = a1 + a2
    assert isinstance(a3, Angle10)
    assert a3 == 100


def test_angle10_plus_angle100() -> None:
    a1 = Angle10(50)
    a2 = Angle100(500)
    a3 = a1 + a2
    assert isinstance(a3, Angle10)
    assert a3 == 100


def test_angle100_plus_angle() -> None:
    a1 = Angle100(500)
    a2 = Angle(5)
    a3 = a1 + a2
    assert isinstance(a3, Angle100)
    assert a3 == 1000


def test_angle100_plus_angle10() -> None:
    a1 = Angle100(500)
    a2 = Angle10(50)
    a3 = a1 + a2
    assert isinstance(a3, Angle100)
    assert a3 == 1000


# endregion Angle + Angle10 + Angle100


# region Angle - Angle10 - Angle100
def test_angle_minus_angle10() -> None:
    a1 = Angle(10)
    a2 = Angle10(50)
    a3 = a1 - a2
    assert isinstance(a3, Angle)
    assert a3 == 5


def test_angle_minus_angle100() -> None:
    a1 = Angle(10)
    a2 = Angle100(500)
    a3 = a1 - a2
    assert isinstance(a3, Angle)
    assert a3 == 5


def test_angle10_minus_angle() -> None:
    a1 = Angle10(100)
    a2 = Angle(5)
    a3 = a1 - a2
    assert isinstance(a3, Angle10)
    assert a3 == 50


def test_angle10_minus_angle100() -> None:
    a1 = Angle10(100)
    a2 = Angle100(500)
    a3 = a1 - a2
    assert isinstance(a3, Angle10)
    assert a3 == 50


def test_angle100_minus_angle() -> None:
    a1 = Angle100(1000)
    a2 = Angle(5)
    a3 = a1 - a2
    assert isinstance(a3, Angle100)
    assert a3 == 500


def test_angle100_minus_angle10() -> None:
    a1 = Angle100(1000)
    a2 = Angle10(50)
    a3 = a1 - a2
    assert isinstance(a3, Angle100)
    assert a3 == 500


# endregion Angle - Angle10 - Angle100


# region Angle * Angle10 * Angle100
def test_angle_multi_angle10() -> None:
    a1 = Angle(10)
    a2 = Angle10(50)
    a3 = a1 * a2
    assert isinstance(a3, Angle)
    assert a3 == 50


def test_angle_multi_angle100() -> None:
    a1 = Angle(10)
    a2 = Angle100(500)
    a3 = a1 * a2
    assert isinstance(a3, Angle)
    assert a3 == 50


def test_angle10_multi_angle() -> None:
    a1 = Angle10(100)
    a2 = Angle(5)
    a3 = a1 * a2
    assert isinstance(a3, Angle10)
    # automatically loops, max is 3600
    # 5000 % 3600 = 1400
    assert a3 == 1400


def test_angle10_multi_angle100() -> None:
    a1 = Angle10(10)
    a2 = Angle100(500)
    a3 = a1 * a2
    assert isinstance(a3, Angle10)
    assert a3 == 500


def test_angle100_multi_angle() -> None:
    a1 = Angle100(1000)
    a2 = Angle(5)  # 5000
    a3 = a1 * a2  # 5000 * 1000 = 5000000
    # automatically loops, max is 36000
    # 5000000 % 36000 = 32_000
    assert isinstance(a3, Angle100)
    assert a3 == 32_000


def test_angle100_multi_angle10() -> None:
    a1 = Angle100(100)
    a2 = Angle10(50)
    a3 = a1 * a2
    assert isinstance(a3, Angle100)
    assert a3 == 14_000


# endregion Angle * Angle10 * Angle100
