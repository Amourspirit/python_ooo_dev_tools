from __future__ import annotations
from typing import Any, TYPE_CHECKING

from ooodev.events.partial.events_partial import EventsPartial
from ooodev.format.inner.partial.factory_styler import FactoryStyler

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst


class DefaultFactoryStyler(FactoryStyler):
    """
    Base Class for Stylers that use a factory.
    """

    def __init__(
        self,
        *,
        factory_name: str,
        component: Any,
        before_event: str,
        after_event: str,
        before_event2: str = "",
        after_event2: str = "",
        lo_inst: LoInst | None = None,
    ) -> None:
        FactoryStyler.__init__(self, factory_name=factory_name, component=component, lo_inst=lo_inst)
        self.after_event_name = after_event
        self.after_event_name2 = after_event2
        self.before_event_name = before_event
        self.before_event_name2 = before_event2
