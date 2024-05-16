from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.adapter.container.index_access_implement import IndexAccessImplement


def test_index_access(loader) -> None:
    lst = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"]
    ia = IndexAccessImplement(lst)
    for i, itm in enumerate(ia):
        assert itm == lst[i]

    reversed_lst = lst[::-1]
    for i, itm in enumerate(reversed(ia)):
        assert itm == reversed_lst[i]

    itm = ia[1]
    assert itm == lst[1]

    slice1 = ia[::-1]
    for i, itm in enumerate(slice1):
        assert itm == reversed_lst[i]

    slice2 = ia[1:5]
    for i, itm in enumerate(slice2):
        assert itm == lst[i + 1]
