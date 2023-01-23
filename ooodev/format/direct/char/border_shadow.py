from __future__ import annotations
from typing import Tuple
from ..structs.shadow import Shadow
from ....meta.static_prop import static_prop

import uno
from ooo.dyn.table.shadow_location import ShadowLocation as ShadowLocation


class BorderShadow(Shadow):
    __EMPTY = None

    def _get_property_name(self) -> str:
        return "CharShadowFormat"

    def _supported_services(self) -> Tuple[str, ...]:
        # will affect apply() on parent class.
        return ("com.sun.star.style.CharacterProperties",)

    @static_prop
    def empty() -> BorderShadow:  # type: ignore[misc]
        """Gets empty Shadow. Static Property. when style is applied it remove any shadow."""
        if BorderShadow._EMPTY is None:
            BorderShadow._EMPTY = BorderShadow(
                location=ShadowLocation.NONE, transparent=False, color=8421504, width=1.76
            )
        return BorderShadow._EMPTY
