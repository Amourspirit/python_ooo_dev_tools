# region Import
from __future__ import annotations
from typing import Any, Tuple, TypeVar

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.format.inner.direct.structs.shadow_struct import ShadowStruct

import uno
from ooo.dyn.table.shadow_location import ShadowLocation as ShadowLocation

# endregion Import

_TShadow = TypeVar(name="_TShadow", bound="Shadow")


class Shadow(ShadowStruct):
    def _get_property_name(self) -> str:
        try:
            return self._property_name
        except AttributeError:
            self._property_name = "CharShadowFormat"
        return self._property_name

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.style.CharacterProperties",
                "com.sun.star.style.CharacterStyle",
            )
        return self._supported_services_values

    def _on_modifying(self, source: Any, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifying(source, event)
