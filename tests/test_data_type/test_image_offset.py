from __future__ import annotations
from ooodev.utils.data_type.image_offset import ImageOffset
import pytest

# ImageOffset implements BaseFloatValue so this test all dunder methods.


@pytest.mark.parametrize("val", [0.0, 0.33, 0.779987, 0.0005, 0.99])
def test_ofsett(val: float) -> None:
    offset = ImageOffset(val)
    assert isinstance(offset, ImageOffset)


@pytest.mark.parametrize("val", [-10.0, -1.22, 400.007, 1.0, 1.001])
def test_offset_error(val: int) -> None:
    with pytest.raises(AssertionError):
        ImageOffset(val)


@pytest.mark.parametrize("val", ["1", 1, -3, None])
def test_offset_error_incorrect_type(val: int) -> None:
    with pytest.raises(TypeError):
        ImageOffset(val)


@pytest.mark.parametrize("start,val", [(0.0, 0.2), (0.224, 0.456), (0.33, 0.055), (0.255, 0.111)])
def test_addition_float(start: float, val: float) -> None:
    a1 = ImageOffset(start)
    a2 = a1 + val
    result = start + val
    assert a2.Value == result


@pytest.mark.parametrize("start,val", [(0.0, 0.2), (0.224, 0.456), (0.33, 0.055), (0.255, 0.111)])
def test_addition_offset_to_offset(start: float, val: float) -> None:
    a1 = ImageOffset(start)
    a2 = ImageOffset(val)
    a3 = a1 + a2
    result = start + val
    assert a3.Value == result


@pytest.mark.parametrize("start,val", [(0.0, 0.2), (0.224, 0.456), (0.33, 0.055), (0.255, 0.111)])
def test_plus_equal(start: float, val: float) -> None:
    a1 = ImageOffset(start)
    a1 += val
    result = start + val
    assert a1.Value == result


def test_angle_init_offset() -> None:
    a1 = ImageOffset(0.1)
    a2 = ImageOffset(float(a1))
    assert a2 == a1


@pytest.mark.parametrize("start,val", [(0.4, 0.2), (0.533, 0.0335589), (0.999, 0.998)])
def test_sub_offset_to_offset(start: float, val: float) -> None:
    a1 = ImageOffset(start)
    a2 = ImageOffset(val)
    a3 = a1 - a2
    result = start - val
    assert a3.Value == result


@pytest.mark.parametrize("start,val", [(0.4, 0.2), (0.533, 0.0335589), (0.999, 0.998)])
def test_minus_equal(start: float, val: float) -> None:
    a1 = ImageOffset(start)
    a1 -= val
    result = start - val
    assert a1.Value == result


def test_sum_offset() -> None:
    a1 = ImageOffset(0.1)
    a2 = ImageOffset(0.2)
    a3 = ImageOffset(0.2)
    a4 = ImageOffset(0.1)
    sums = sum([a1, a2, a3, a4])
    assert isinstance(sums, ImageOffset)
    assert sums.Value == 0.6


@pytest.mark.parametrize("start,val", [(0.08, 0.2), (0.2, 0.008), (0.5, 0.499), (0.23, 0.3)])
def test_mul_offset_to_offset(start: float, val: float) -> None:
    a1 = ImageOffset(start)
    a2 = ImageOffset(val)
    a3 = a1 * a2
    result = start * val
    assert a3.Value == result


@pytest.mark.parametrize("start,val", [(0.1, 0.9), (0.02, 0.3)])
def test_lt_offset(start: float, val: float) -> None:
    a1 = ImageOffset(start)
    a2 = ImageOffset(val)
    assert a1 < a2
    gt = a1 > a2
    assert gt == False


@pytest.mark.parametrize("start,val", [(0.11, 0.11), (0.3, 0.4), (0.003, 0.003), (0.77, 0.9)])
def test_le_offset(start: float, val: float) -> None:
    a1 = ImageOffset(start)
    a2 = ImageOffset(val)
    assert a1 <= a2
    gt = a1 > a2
    assert gt == False


@pytest.mark.parametrize("start,val", [(0.9, 0.1), (0.3, 0.02)])
def test_gt_offset(start: float, val: float) -> None:
    a1 = ImageOffset(start)
    a2 = ImageOffset(val)
    assert a1 > a2
    lt = a1 < a2
    assert lt == False


@pytest.mark.parametrize("start,val", [(0.33, 0.33), (0.4, 0.2), (0.0, 0.0), (0.43, 0.42)])
def test_ge_offset(start: int, val: int) -> None:
    a1 = ImageOffset(start)
    a2 = ImageOffset(val)
    assert a1 >= a2
    lt = a1 < a2
    assert lt == False
