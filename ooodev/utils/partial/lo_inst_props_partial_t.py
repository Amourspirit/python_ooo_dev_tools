from __future__ import annotations
from typing import TYPE_CHECKING
from ooodev.loader.inst.lo_inst import LoInst

if TYPE_CHECKING:
    from typing_extensions import Protocol
else:
    Protocol = object


class LoInstPropsPartialT(Protocol):
    """Type for LoInstPropsPartial"""

    @property
    def lo_inst(self) -> LoInst:
        """Lo Instance"""
        ...
