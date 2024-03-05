from __future__ import annotations
from typing import Any, TYPE_CHECKING

from ooodev.format.inner.partial.default_factor_styler import DefaultFactoryStyler
from ooodev.format.inner.style_factory import area_color_factory
from ooodev.utils import color as mColor
from ooodev.events.partial.events_partial import EventsPartial

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.format.proto.area.fill_color_t import FillColorT
else:
    FillColorT = Any


class FillColorPartial:
    """
    Partial class for FillColor.
    """

    def __init__(self, factory_name: str, component: Any, lo_inst: LoInst | None = None) -> None:
        self.__styler = DefaultFactoryStyler(
            factory_name=factory_name,
            component=component,
            before_event="before_style_area_color",
            after_event="after_style_area_color",
            lo_inst=lo_inst,
        )
        if isinstance(self, EventsPartial):
            self.__styler.add_event_observers(self.event_observer)

    def style_area_color(self, color: mColor.Color = mColor.StandardColor.AUTO_COLOR) -> FillColorT | None:
        """
        Style Area Color.

        Args:
            color (:py:data:`~.utils.color.Color`, optional): FillColor Color. Defaults to ``StandardColor.AUTO_COLOR``.

        Raises:
            CancelEventError: If the event ``before_style_area_color`` is cancelled and not handled.

        Returns:
            FillColorT | None: FillColor instance or ``None`` if cancelled.
        """

        return self.__styler.style(factory=area_color_factory, color=color)

    def style_area_color_get(self) -> FillColorT | None:
        """
        Gets the Area Color Style.

        Raises:
            CancelEventError: If the event ``before_style_area_color_get`` is cancelled and not handled.

        Returns:
            FillColorT | None: Area color style or ``None`` if cancelled.
        """
        return self.__styler.style_get(factory=area_color_factory)
