from __future__ import annotations
from typing import Any

from .....events.args.cancel_event_args import CancelEventArgs
from ...structs.shadow_struct import ShadowStruct
from ooo.dyn.table.shadow_location import ShadowLocation as ShadowLocation


class Shadow(ShadowStruct):
    def _on_modifing(self, source: Any, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifing(source, event)
