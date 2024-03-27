from __future__ import annotations
from typing import Any
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.loader import lo as mLo


class DispatchContext:
    """
    Context manager for toggling the dispatch command.
    This manager sends the command to the LibreOffice instance on enter and exit.

    The dispatch command is used to send commands to the LibreOffice instance.

    Example:
        .. code-block:: python

            with DispatchContext("SwitchControlDesignMode"):
                ctl = cell.control.insert_control_currency_field(
                    min_value=0, spin_button=True
                )
                ctl.model.Repeat = True
                ctl.value = cell.get_num()

    See Also:
        - :py:meth:`ooodev.loader.Lo.dispatch_cmd`

    .. versionadded:: 0.38.1
    """

    def __init__(self, cmd: str, inst: LoInst | None = None, **kwargs: Any) -> None:
        """
        Constructor for LoContext

        Args:
            cmd (str): The command to send to the LibreOffice instance.
            inst (LoInst, optional): The instance of that is used to send dispatch. This is generally used when working with multiple instances of LibreOffice.
            kwargs (Any, optional): Any additional arguments to send with the command.
        """
        if inst is None:
            inst = mLo.Lo.current_lo
        self._inst = inst
        self._cmd = cmd
        self._kwargs = kwargs

    def __enter__(self) -> None:
        self._inst.dispatch_cmd(self._cmd, **self._kwargs)
        return None

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        self._inst.dispatch_cmd(self._cmd, **self._kwargs)
        return None
