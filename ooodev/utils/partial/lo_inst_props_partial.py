from __future__ import annotations
from ooodev.utils.inst.lo.lo_inst import LoInst


class LoInstPropsPartial:
    """Partial Used implementing LoInst. Property"""

    def __init__(self, lo_inst: LoInst) -> None:
        """
        Constructor.

        Args:
            lo_inst (LoInst, optional): Lo instance.
        """
        self.__lo_inst = lo_inst

    @property
    def lo_inst(self) -> LoInst:
        """Lo Instance"""
        return self.__lo_inst
