from __future__ import annotations
from ooodev.utils.data_type.intensity import Intensity
import pytest

# intensity implements BaseIntValue so this test all dunder methods.


def test_intensity() -> None:
    for i in range(101):
        intensity = Intensity(i)
        assert isinstance(intensity, Intensity)


@pytest.mark.parametrize("val", ["1", 1.2, -3.3])
def test_intensity_error_incorrect_type(val: int) -> None:
    with pytest.raises(TypeError):
        Intensity(val)


@pytest.mark.parametrize("start,val", [(0, 5), (1, 3), (33, 55)])
def test_addition_int(start: int, val: int) -> None:
    a1 = Intensity(start)
    a2 = a1 + val
    assert a2.value == start + val


@pytest.mark.parametrize("start,val", [(0, 0), (0, 5), (1, 3), (33, 55)])
def test_addition_intensity_to_intensity(start: int, val: int) -> None:
    a1 = Intensity(start)
    a2 = Intensity(val)
    a3 = a1 + a2
    assert a3.value == start + val


@pytest.mark.parametrize("start,val", [(0, 0), (0, 5), (1, 3), (33, 55)])
def test_addition_plus_equal(start: int, val: int) -> None:
    a1 = Intensity(start)
    a1 += val
    result = start + val
    assert a1.value == result


def test_intensity_init_intensity() -> None:
    a1 = Intensity(2)
    a2 = Intensity(int(a1))
    assert a2 == a1


@pytest.mark.parametrize("start,val", [(0, 0), (5, 0), (3, 1), (55, 22)])
def test_sub_intensity_to_intensity(start: int, val: int) -> None:
    a1 = Intensity(start)
    a2 = Intensity(val)
    a3 = a1 - a2
    assert a3.value == start - val


@pytest.mark.parametrize("start,val", [(0, 0), (5, 0), (3, 1), (55, 22)])
def test_minus_equal(start: int, val: int) -> None:
    a1 = Intensity(start)
    a1 -= val
    result = start - val
    assert a1.value == result


def test_sum_intensity() -> None:
    a1 = Intensity(5)
    a2 = Intensity(10)
    a3 = Intensity(15)
    a4 = Intensity(20)
    sums = sum([a1, a2, a3, a4])
    assert isinstance(sums, Intensity)
    assert sums.value == 50


@pytest.mark.parametrize("start,val", [(0, 0), (5, 0), (3, 3), (33, 3)])
def test_mul_intensity_to_intensity(start: int, val: int) -> None:
    a1 = Intensity(start)
    a2 = Intensity(val)
    a3 = a1 * a2
    result = start * val
    assert a3.value == result


@pytest.mark.parametrize("start,val", [(5, 2), (100, 25), (99, 3)])
def test_div_intensity_to_intensity(start: int, val: int) -> None:
    a1 = Intensity(start)
    a2 = Intensity(val)
    a3 = a1 / a2
    result = round(start / val)
    assert a3.value == result


@pytest.mark.parametrize("start,val", [(0, 1), (2, 3), (3, 77), (55, 100), (88, 89)])
def test_lt_intensity(start: int, val: int) -> None:
    a1 = Intensity(start)
    a2 = Intensity(val)
    assert a1 < a2
    gt = a1 > a2
    assert gt == False


@pytest.mark.parametrize("start,val", [(0, 1), (2, 3), (5, 5), (55, 100), (100, 100)])
def test_le_intensity(start: int, val: int) -> None:
    a1 = Intensity(start)
    a2 = Intensity(val)
    assert a1 <= a2
    gt = a1 > a2
    assert gt == False


@pytest.mark.parametrize("start,val", [(1, 0), (3, 2), (77, 3), (100, 10), (88, 3)])
def test_gt_intensity(start: int, val: int) -> None:
    a1 = Intensity(start)
    a2 = Intensity(val)
    assert a1 > a2
    lt = a1 < a2
    assert lt == False


@pytest.mark.parametrize("start,val", [(1, 0), (3, 2), (5, 5), (100, 44), (100, 100), (0, 0)])
def test_ge_intensity(start: int, val: int) -> None:
    a1 = Intensity(start)
    a2 = Intensity(val)
    assert a1 >= a2
    lt = a1 < a2
    assert lt == False


@pytest.mark.parametrize(("val", "expected"), [(360, 0), (-3, 357), (-181, 179), (2500, 340), (-2500, 20)])
def test_intensity_weird(val: int, expected: int) -> None:
    a1 = Intensity(val)
    assert a1.value == expected


@pytest.mark.parametrize(("val", "subval", "expected"), [(1, 100, 99), (2, 3, 1)])
def test_intensity_rsub(val: int, subval: int, expected: int) -> None:
    # test __rsub__
    # Test when a number subtracts an Intensity
    a1 = Intensity(val)
    result = subval - a1
    assert result == Intensity(expected)


@pytest.mark.parametrize("start", [1, 3, 5, 100, 0])
def test_eq_intensity(start: int) -> None:
    a1 = Intensity(start)
    a2 = Intensity(start)
    assert a1 == a2
    assert a2 == a1
    assert a1 == start
    assert start == a1
