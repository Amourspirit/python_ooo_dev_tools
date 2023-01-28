from __future__ import annotations

from .....meta.static_prop import static_prop
from ...structs.shadow_struct import ShadowStruct
from ooo.dyn.table.shadow_location import ShadowLocation as ShadowLocation


class Shadow(ShadowStruct):
    _EMPTY = None

    @static_prop
    def empty() -> ShadowStruct:  # type: ignore[misc]
        """Gets empty Shadow. Static Property. when style is applied it remove any shadow."""
        if Shadow._EMPTY is None:
            Shadow._EMPTY = Shadow(location=ShadowLocation.NONE, transparent=False, color=8421504, width=1.76)
        return Shadow._EMPTY
