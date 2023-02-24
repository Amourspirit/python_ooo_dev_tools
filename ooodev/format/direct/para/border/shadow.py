from __future__ import annotations
from typing import Tuple

from .....events.args.cancel_event_args import CancelEventArgs
from ...structs.shadow_struct import ShadowStruct
from .....meta.static_prop import static_prop

import uno
from ooo.dyn.table.shadow_location import ShadowLocation as ShadowLocation


class Shadow(ShadowStruct):
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

    @static_prop
    def empty() -> Shadow:  # type: ignore[misc]
        """Gets empty Shadow. Static Property. when style is applied it remove any shadow."""
        try:
            return Shadow._EMPTY_INST
        except AttributeError:
            Shadow._EMPTY_INST = Shadow(location=ShadowLocation.NONE, transparent=False, color=8421504, width=1.76)
            Shadow._EMPTY_INST._is_default_inst = True
        return Shadow._EMPTY_INST
