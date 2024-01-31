from __future__ import annotations
from typing import TYPE_CHECKING
from ooodev.utils.inst.lo.lo_inst import LoInst

if TYPE_CHECKING:
    try:
        from typing import Protocol
    except ImportError:
        from typing_extensions import Protocol
else:
    Protocol = object


class LoInstPropsPartialT(Protocol):
    """Type for LoInstPropsPartial"""

    @property
    def lo_inst(self) -> LoInst:
        """Lo Instance"""
        ...
