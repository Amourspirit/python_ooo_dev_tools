from typing import cast
import pytest

if __name__ == "__main__":
    pytest.main([__file__])


def test_row_math() -> None:
    from ooodev.utils.data_type.row_obj import RowObj

    r1 = RowObj.from_int(1)
    assert r1.value == 1
    assert r1.index == 0

    r2 = RowObj(2)
    assert r2 > r1
    assert r1 < r2
    assert r2 <= r2
    assert r2 >= r2
    assert r2 >= 2
    assert 2 <= r2
    assert r2 == 2
    assert 2 == r2

    r3 = r1 + r2
    assert r3.value == 3

    r3 = r1 + 2
    assert r3.value == 3

    r2 = r3 - r1
    assert r2.value == 2

    r2 = r3 - 1
    assert r2.value == 2

    r2 = 5 - r3
    assert r2.value == 2


def test_row_math_errors() -> None:
    from ooodev.utils.data_type.row_obj import RowObj

    r1 = RowObj.from_int(1)
    r3 = RowObj(3)
    with pytest.raises(IndexError):
        _ = r1 - r3
    with pytest.raises(IndexError):
        _ = r1 - 1

    with pytest.raises(IndexError):
        _ = r1 - 3

    with pytest.raises(IndexError):
        _ = r1 + -3

    with pytest.raises(IndexError):
        _ = r3 - r3

    with pytest.raises(IndexError):
        _ = r3 - 3

    with pytest.raises(IndexError):
        _ = 0 - r1

    with pytest.raises(IndexError):
        _ = 1 - r1

    with pytest.raises(IndexError):
        _ = 3 - r3

    with pytest.raises(IndexError):
        _ = -4 + r3
