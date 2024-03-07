from __future__ import annotations
from typing import Any, Dict, TYPE_CHECKING

from ooodev.format.inner.style_factory import area_transparency_transparency_factory
from ooodev.format.inner.partial.default_factor_styler import DefaultFactoryStyler
from ooodev.events.partial.events_partial import EventsPartial

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.utils.data_type.intensity import Intensity
    from ooodev.format.proto.area.transparency.transparency_t import TransparencyT
else:
    LoInst = Any
    TransparencyT = Any
    Intensity = Any


class TransparencyPartial:
    """
    Partial class for FillColor.
    """

    def __init__(self, factory_name: str, component: Any, lo_inst: LoInst | None = None) -> None:
        self.__styler = DefaultFactoryStyler(
            factory_name=factory_name,
            component=component,
            before_event="before_style_area_transparency_transparency",
            after_event="after_style_area_transparency_transparency",
            lo_inst=lo_inst,
        )
        if isinstance(self, EventsPartial):
            self.__styler.add_event_observers(self.event_observer)

    def style_area_transparency_transparency(self, value: Intensity | int = 0) -> TransparencyT | None:
        """
        Style Area Color.

        Args:
            value (Intensity, int, optional): Specifies the transparency value from ``0`` to ``100``.

        Raises:
            CancelEventError: If the event ``before_style_area_transparency_transparency`` is cancelled and not handled.

        Returns:
            TransparencyT | None: FillColor instance or ``None`` if cancelled.

        Hint:
            - The value of ``0`` is fully opaque.
            - The value of ``100`` is fully transparent.
            - ``Intensity`` can be imported from ``ooodev.utils.data_type.intensity``
        """
        return self.__styler.style(factory=area_transparency_transparency_factory, value=value)

    def style_area_transparency_transparency_get(self) -> TransparencyT | None:
        """
        Gets the Area Transparency Style.

        Raises:
            CancelEventError: If the event ``before_style_area_transparency_transparency_get`` is cancelled and not handled.

        Returns:
            TransparencyT | None: Area transparency style or ``None`` if cancelled.
        """
        return self.__styler.style_get(factory=area_transparency_transparency_factory)
