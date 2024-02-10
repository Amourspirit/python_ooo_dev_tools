from __future__ import annotations
from typing import Any, TYPE_CHECKING

from ooodev.loader import lo as mLo
from ooodev.events.partial.events_partial import EventsPartial

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
else:
    LoInst = Any


class FactoryNameBase(EventsPartial):
    """
    Base Class for Factory Name
    """

    def __init__(self, factory_name: str, component: Any, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        EventsPartial.__init__(self)
        self._lo_inst = lo_inst
        self._factory_name = factory_name
        self._component = component
        self._before_event_name = f"before_{factory_name}"
        self._after_event_name = f"after_{factory_name}"
        self._before_event_name2 = f"before_{factory_name}2"
        self._after_event_name2 = f"after_{factory_name}2"

    @property
    def lo_inst(self) -> LoInst:
        return self._lo_inst

    @lo_inst.setter
    def lo_inst(self, value: LoInst) -> None:
        self._lo_inst = value

    @property
    def factory_name(self) -> str:
        return self._factory_name

    @factory_name.setter
    def factory_name(self, value: str) -> None:
        self._factory_name = value

    @property
    def component(self) -> Any:
        return self._component

    @component.setter
    def component(self, value: Any) -> None:
        self._component = value

    @property
    def after_event_name(self) -> str:
        return self._after_event_name

    @after_event_name.setter
    def after_event_name(self, value: str) -> None:
        self._after_event_name = value

    @property
    def before_event_name(self) -> str:
        return self._before_event_name

    @before_event_name.setter
    def before_event_name(self, value: str) -> None:
        self._before_event_name = value

    @property
    def after_event_name2(self) -> str:
        return self._after_event_name2

    @after_event_name2.setter
    def after_event_name2(self, value: str) -> None:
        self._after_event_name2 = value

    @property
    def before_event_name2(self) -> str:
        return self._before_event_name2

    @before_event_name2.setter
    def before_event_name2(self, value: str) -> None:
        self._before_event_name2 = value
