from __future__ import annotations
from ooodev.utils.inst.lo.lo_inst import LoInst
from ooodev.utils import lo as mLo


class LoContext:
    def __init__(self, inst: LoInst) -> None:
        self._current = mLo.Lo.current_lo
        self._inst = inst

    def __enter__(self) -> LoInst:
        # switch the context
        # only switch if not default
        if self._current is self._inst:
            return self._inst
        if self._inst.is_default:
            return self._inst
        mLo.Lo._lo_inst = self._inst
        return self._inst

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        if self._current is self._inst:
            return None
        if self._inst.is_default:
            return None
        mLo.Lo._lo_inst = self._current
        return None
