from typing import cast
import pytest

if __name__ == "__main__":
    pytest.main([__file__])


def test_col_math() -> None:
    from ooodev.utils.data_type.col_obj import ColObj

    c1 = ColObj.from_int(1)
    assert c1.index == 0
    assert c1.value == "A"

    c1 = ColObj.from_int(0, True)
    assert c1.index == 0
    assert c1.value == "A"
    assert c1 == 1
    assert 1 == c1

    c1 = ColObj("A")
    assert c1.index == 0
    assert c1.value == "A"
    assert c1 == "A"

    c2 = ColObj("b")
    assert c2.index == 1
    assert c2.value == "B"
    assert c2 == "b"

    assert c2 > c1
    assert c1 < c2

    assert c2 == "B"
    assert "B" == c2

    assert c1 == "a"
    assert "a" == c1

    assert c2 > "a"
    assert "A" < c2

    assert c2 >= "b"
    assert c2 >= "a"

    assert "c" > c2
    assert c2 < "c"

    c3 = c2 + 1
    assert c3.index == 2
    assert c3.value == "C"
    assert c3 == "C"

    c5 = ColObj("E")
    assert c5.index == 4

    c10 = c1 + 9
    assert c10.value == "J"
    assert c10.index == 9

    c11 = c2 + 9
    assert c11.value == "K"
    assert c11.index == 10
    assert c11 > c10

    # I is adding 9
    # b (2) + I (9) = K (11)
    c11 = c2 + "i"
    assert c11.value == "K"
    assert c11.index == 10
    assert c11 > c10

    c11 = "i" + c2
    assert c11.value == "K"
    assert c11.index == 10
    assert c11 > c10
    assert c11 < "O"

    c9 = c11 - 2
    assert c9.value == "I"
    assert c9.index == 8
    assert c9 < c11
    assert c9 < "K"

    c9 = c11 - c2
    assert c9.value == "I"
    assert c9.index == 8
    assert c9 < c11
    assert c9 < "K"

    c9 = c11 - "B"
    assert c9.value == "I"
    assert c9.index == 8
    assert c9 < c11
    assert c9 < "K"

    c2 = 5 - c3
    assert c2.index == 1

    c2 = "E" - c3
    assert c2.index == 1

    assert c1.index == 0
    assert c2.index == 1
    assert c3.index == 2
    cc = c1 + c2
    assert cc.index == 2
    assert c1 + c2 == c3

    c_sum = sum([c1, c2, c3])
    assert c_sum.index == 5

    c_sum = sum([c1, 2, 3])
    assert c_sum.index == 5

    c_sum = sum([c1, "b", "c"])
    assert c_sum.index == 5

    c_sum = sum([c1, "b", 3])
    assert c_sum.index == 5

    c6 = c2 * 3
    assert c6.index == 5

    c60 = c6 * 10
    assert c60.index == 59

    c6 = c60 / 10
    assert c6.index == 5

    c10 = (c6 - 1) + c5
    assert c10.index == 9

    c5 = 50 / c10
    assert c5.index == 4

    c3 = c6 / "B"  # divide by 2
    assert c3.index == 2

    c12 = c3 * "D"  # muliply by 4
    assert c12.index == 11

    c4 = "T" / c5  # T is 20
    assert c4.index == 3


def test_col_math_errors() -> None:
    from ooodev.utils.data_type.col_obj import ColObj

    c1 = ColObj("A")
    c3 = ColObj("C")

    with pytest.raises(IndexError):
        _ = c1 - c3

    with pytest.raises(IndexError):
        _ = c1 - 3

    with pytest.raises(IndexError):
        _ = c1 - 1

    with pytest.raises(IndexError):
        _ = c1 - "C"

    with pytest.raises(IndexError):
        _ = -10 + c3

    with pytest.raises(IndexError):
        _ = c3 / 5

    with pytest.raises(IndexError):
        _ = c3 / "E"

    with pytest.raises(IndexError):
        _ = c3 / 0
    
    with pytest.raises(IndexError):
        _ = "B" / c3


def test_col_next_prev() -> None:
    from ooodev.utils.data_type.col_obj import ColObj

    c1 = ColObj("A")
    c2 = c1.next
    assert c2.value == "B"
    assert c2.prev == c1

    c1 = None
    c1 = c2.prev
    assert c1.value == "A"

    c5 = c1.next.next.next.next
    assert c5.value == "E"
    assert c5 == 5

    assert c5.prev == "D"
    assert c5.prev == 4
    assert c5.prev.index == 3

    assert c5.next == "F"
    assert c5.next == 6
    assert c5.next.index == 5

    assert c5.next > c5
    assert c5 < c5.next
    assert c5.prev.prev < c5.prev
    assert c5.next.next > c5.next

    col = ColObj("A")
    for _ in range(10):
        col = col.next
    assert col.value == "K"

    for _ in range(10):
        col = col.prev
    assert col.value == "A"


def test_col_prev_error() -> None:
    from ooodev.utils.data_type.col_obj import ColObj

    c1 = ColObj("A")
    c2 = c1.next
    with pytest.raises(IndexError):
        _ = c1.prev

    with pytest.raises(IndexError):
        _ = c2.prev.prev

    with pytest.raises(IndexError):
        _ = c2.prev.prev.prev.prev
