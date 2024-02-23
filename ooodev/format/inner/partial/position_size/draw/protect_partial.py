from __future__ import annotations
from typing import Any, TYPE_CHECKING

from ooodev.events.partial.events_partial import EventsPartial
from ooodev.format.inner.partial.factory_styler import FactoryStyler
from ooodev.format.inner.style_factory import draw_position_size_protect_factory
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.format.proto.draw.position_size.protect_t import ProtectT
else:
    ProtectT = Any


class ProtectPartial:
    """
    Partial class for Size.
    """

    def __init__(self, factory_name: str, component: Any, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self.__styler = FactoryStyler(factory_name=factory_name, component=component, lo_inst=lo_inst)
        if isinstance(self, EventsPartial):
            self.__styler.add_event_observers(self.event_observer)
        self.__styler.after_event_name = "after_style_protect"
        self.__styler.before_event_name = "before_style_protect"

    def style_protect(self, position: bool | None = None, size: bool | None = None) -> ProtectT | None:
        """
        Style Area Color.

        Args:
            position (bool, optional): Specifies position protection.
            size (bool, optional): Specifies size protection.

        Raises:
            CancelEventError: If the event ``before_style_protect`` is cancelled and not handled.

        Returns:
            ProtectT | None: Position instance or ``None`` if cancelled.
        """
        return self.__styler.style(factory=draw_position_size_protect_factory, position=position, size=size)

    def style_protect_get(self) -> ProtectT | None:
        """
        Gets the Protect Style.

        Raises:
            CancelEventError: If the event ``before_style_protect_get`` is cancelled and not handled.

        Returns:
            ProtectT | None: Protect style or ``None`` if cancelled.
        """
        return self.__styler.style_get(factory=draw_position_size_protect_factory)
