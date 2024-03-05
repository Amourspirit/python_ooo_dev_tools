from __future__ import annotations
from typing import Any, TYPE_CHECKING

from ooodev.events.partial.events_partial import EventsPartial
from ooodev.format.inner.partial.factory_styler import FactoryStyler
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst


class FactoryStylerBase:
    """
    Base Class for Stylers that use a factory.
    """

    def __init__(
        self, factory_name: str, component: Any, before_event: str, after_event: str, lo_inst: LoInst | None = None
    ) -> None:
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        styler = FactoryStyler(factory_name=factory_name, component=component, lo_inst=lo_inst)
        if isinstance(self, EventsPartial):
            styler.add_event_observers(self.event_observer)
        styler.after_event_name = after_event
        styler.before_event_name = before_event
        self.__styler = styler

    def _get_styler(self) -> FactoryStyler:
        return self.__styler
