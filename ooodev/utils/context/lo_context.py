from __future__ import annotations
from time import sleep, time
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.loader import lo as mLo


class LoContext:
    """
    LoContext is a context manager for switching the current LibreOffice context to a different instance.

    This allows for multiple instances of LibreOffice Documents to be used at the same time.
    """

    def __init__(self, inst: LoInst, timeout_ms: float = 30_000) -> None:
        """
        Constructor for LoContext

        Args:
            inst (LoInst): The instance of LO Context to switch to.
            timeout_ms (float, optional): Timeout in milliseconds. Set to ``0`` for no timeout.  Defaults to ``30_000`` ( ``30`` seconds).
        """
        self._current = mLo.Lo.current_lo
        self._inst = inst
        self._timeout = timeout_ms  # timeout_ms / 1000

    def __enter__(self) -> LoInst:
        # switch the context
        # only switch if not default
        if self._current is self._inst:
            return self._inst
        if self._inst.is_default:
            return self._inst
        # if any other instance of context manager has locked the context
        # wait for it to release.
        if self._timeout > 0:
            end_time = time() + (self._timeout / 1000)
            while mLo.Lo._locked:
                sleep(0.1)
                if end_time > time():
                    raise TimeoutError("Timeout waiting for LoContext to unlock.")
        else:
            while mLo.Lo._locked:
                sleep(0.1)
        mLo.Lo._lo_inst = self._inst
        mLo.Lo._locked = True
        return self._inst

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        if self._current is self._inst:
            return None
        if self._inst.is_default:
            return None
        mLo.Lo._lo_inst = self._current
        mLo.Lo._locked = False
        return None
