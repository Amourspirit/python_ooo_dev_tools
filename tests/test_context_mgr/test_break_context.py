from ooodev.wrapper.break_context import BreakContext

class CtxTest:

    def __init__(self, val:int) -> None:

        self._value = val

    def __enter__(self) -> int:
        return self._value

    def __exit__(self, exc_type, exc_val, exc_tb):
        self._value = 0

def test_break_with() -> None:
    with BreakContext(CtxTest(10)) as val:
        assert val == 10
    assert val == 10

def test_break_err() -> None:
    myval = 25
    with BreakContext(CtxTest(10)) as val:
        if val == 10:
            raise BreakContext.Break
        myal = val
    assert myval == 25