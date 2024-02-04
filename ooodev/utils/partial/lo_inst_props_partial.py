from __future__ import annotations
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst


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
