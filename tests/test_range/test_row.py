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


def test_row_next_prev() -> None:
    from ooodev.utils.data_type.row_obj import RowObj

    # next and prev use weak ref internaly
    r1 = RowObj(1)
    r2 = r1.next
    assert r2.value == 2
    r3_1 = r2.next
    assert r3_1.value == 3

    r2 = None

    r2 = r1.next
    assert r2.value == 2

    r1 = None
    r1 = r2.prev
    assert r1.value == 1
    r5 = r1.next.next.next.next
    assert r5.value == 5

    assert r5.prev == 4
    assert r5.next == 6
    assert r5.next > r5.prev
    assert r5.prev < r5.next
    assert r5.next < r5.next.next
    assert r5.next.next > r5
    assert r5.next.next > r5.next

    row = RowObj(1)
    for _ in range(10):
        row = row.next
    assert row.value == 11

    for _ in range(10):
        row = row.prev
    assert row.value == 1


def test_row_prev_error() -> None:
    from ooodev.utils.data_type.row_obj import RowObj

    r1 = RowObj(1)
    r2 = r1.next
    with pytest.raises(IndexError):
        _ = r1.prev

    with pytest.raises(IndexError):
        _ = r2.prev.prev

    with pytest.raises(IndexError):
        _ = r2.prev.prev.prev.prev
