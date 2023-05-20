from __future__ import annotations
from typing import Tuple

import uno
from ooo.dyn.table.shadow_location import ShadowLocation as ShadowLocation

from ooodev.format.inner.direct.structs.shadow_struct import ShadowStruct


class Shadow(ShadowStruct):
    """Shadow format for paragraph border"""

    def _get_property_name(self) -> str:
        try:
            return self._property_name
        except AttributeError:
            self._property_name = "ParaShadowFormat"
        return self._property_name

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.style.CharacterProperties",
                "com.sun.star.style.ParagraphStyle",
            )
        return self._supported_services_values
