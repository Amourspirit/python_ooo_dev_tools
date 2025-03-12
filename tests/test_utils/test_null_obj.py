from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])


def test_null_obj() -> None:
    from ooodev.utils.gen_util import NULL_OBJ

    # Test that NULL_OBJ evaluates to False in boolean context
    assert bool(NULL_OBJ) is False

    # Test in if statement context
    if NULL_OBJ:
        pytest.fail("NULL_OBJ should evaluate to False")

    # Test that NULL_OBJ is not None
    assert NULL_OBJ is not None

    # Test that multiple NULL_OBJ references point to the same instance
    from ooodev.utils.gen_util import NULL_OBJ as NULL_OBJ2

    assert NULL_OBJ is NULL_OBJ2

    x = NULL_OBJ
    if x:
        pytest.fail("NULL_OBJ should evaluate to False")
