from __future__ import annotations
from ooodev.utils.data_type.angle import Angle
import pytest

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


@pytest.mark.parametrize("start,val", [(0, 0), (0, 5), (1, 3), (33, 55), (255, 100)])
def test_addition_angle_to_angle(start: int, val: int) -> None:
    a1 = Angle(start)
    a2 = Angle(val)
    a3 = a1 + a2
    assert a3.value == start + val


@pytest.mark.parametrize("start,val", [(0, 0), (0, 5), (1, 3), (33, 55), (255, 100)])
def test_addition_plus_equal(start: int, val: int) -> None:
    a1 = Angle(start)
    a1 += val
    result = start + val
    assert a1.value == result


def test_angle_init_angle() -> None:
    a1 = Angle(2)
    a2 = Angle(int(a1))
    assert a2 == a1


@pytest.mark.parametrize("start,val", [(0, 0), (5, 0), (3, 1), (55, 22), (255, 100)])
def test_sub_angle_to_angle(start: int, val: int) -> None:
    a1 = Angle(start)
    a2 = Angle(val)
    a3 = a1 - a2
    assert a3.value == start - val


@pytest.mark.parametrize("start,val", [(0, 0), (5, 0), (3, 1), (55, 22), (255, 100)])
def test_minus_equal(start: int, val: int) -> None:
    a1 = Angle(start)
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


@pytest.mark.parametrize("start,val", [(0, 0), (5, 0), (3, 3), (55, 3)])
def test_mul_angle_to_angle(start: int, val: int) -> None:
    a1 = Angle(start)
    a2 = Angle(val)
    a3 = a1 * a2
    result = start * val
    assert a3.value == result


@pytest.mark.parametrize("start,val", [(5, 2), (300, 45), (355, 3)])
def test_div_angle_to_angle(start: int, val: int) -> None:
    a1 = Angle(start)
    a2 = Angle(val)
    a3 = a1 / a2
    result = round(start / val)
    assert a3.value == result


@pytest.mark.parametrize("start,val", [(0, 1), (2, 3), (3, 77), (55, 198), (100, 359)])
def test_lt_angle(start: int, val: int) -> None:
    a1 = Angle(start)
    a2 = Angle(val)
    assert a1 < a2
    gt = a1 > a2
    assert gt == False


@pytest.mark.parametrize("start,val", [(0, 1), (2, 3), (5, 5), (55, 198), (360, 360)])
def test_le_angle(start: int, val: int) -> None:
    a1 = Angle(start)
    a2 = Angle(val)
    assert a1 <= a2
    gt = a1 > a2
    assert gt == False


@pytest.mark.parametrize("start,val", [(1, 0), (3, 2), (77, 3), (198, 10), (359, 3)])
def test_gt_angle(start: int, val: int) -> None:
    a1 = Angle(start)
    a2 = Angle(val)
    assert a1 > a2
    lt = a1 < a2
    assert lt == False


@pytest.mark.parametrize("start,val", [(1, 0), (3, 2), (5, 5), (198, 44), (360, 360)])
def test_ge_angle(start: int, val: int) -> None:
    a1 = Angle(start)
    a2 = Angle(val)
    assert a1 >= a2
    lt = a1 < a2
    assert lt == False


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
