from __future__ import annotations
from typing import Any, Type, TypeVar

from .....events.args.cancel_event_args import CancelEventArgs
from .....meta.class_property_readonly import ClassPropertyReadonly
from ...structs.shadow_struct import ShadowStruct
from ooo.dyn.table.shadow_location import ShadowLocation as ShadowLocation

_TShadow = TypeVar(name="_TShadow", bound="Shadow")


class Shadow(ShadowStruct):
    def _on_modifing(self, source: Any, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifing(source, event)

    @ClassPropertyReadonly
    @classmethod
    def empty(cls: Type[_TShadow]) -> _TShadow:  # type: ignore[misc]
        """Gets empty Shadow. Static Property. when style is applied it remove any shadow."""
        try:
            return cls._EMPTY_INST
        except AttributeError:
            cls._EMPTY_INST = cls(location=ShadowLocation.NONE, transparent=False, color=8421504, width=1.76)
            cls._EMPTY_INST._is_default_inst = True
        return cls._EMPTY_INST
