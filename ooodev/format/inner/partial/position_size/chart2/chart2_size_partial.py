from __future__ import annotations
from typing import Any, TYPE_CHECKING

from ooodev.format.inner.partial.default_factor_styler import DefaultFactoryStyler
from ooodev.format.inner.style_factory import chart2_position_size_size_factory
from ooodev.events.partial.events_partial import EventsPartial

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.format.proto.chart2.position_size.size_t import SizeT
    from ooodev.units.unit_obj import UnitT
else:
    SizeT = Any
    UnitT = Any


class Chart2SizePartial:
    """
    Partial class for Chart Size.
    """

    def __init__(self, factory_name: str, component: Any, lo_inst: LoInst | None = None) -> None:
        self.__styler = DefaultFactoryStyler(
            factory_name=factory_name,
            component=component,
            before_event="before_style_size",
            after_event="after_style_size",
            lo_inst=lo_inst,
        )
        if isinstance(self, EventsPartial):
            self.__styler.add_event_observers(self.event_observer)

    def style_size(self, width: float | UnitT, height: float | UnitT) -> SizeT | None:
        """
        Style Area Color.

        Args:
            width (float | UnitT): Specifies the width of the shape (in ``mm`` units) or :ref:`proto_unit_obj`.
            height (float | UnitT): Specifies the height of the shape (in ``mm`` units) or :ref:`proto_unit_obj`.

        Raises:
            CancelEventError: If the event ``before_style_size`` is cancelled and not handled.

        Returns:
            SizeT | None: Position instance or ``None`` if cancelled.
        """
        return self.__styler.style(factory=chart2_position_size_size_factory, width=width, height=height)

    def style_size_get(self) -> SizeT | None:
        """
        Gets the Size Style.

        Raises:
            CancelEventError: If the event ``before_style_size_get`` is cancelled and not handled.

        Returns:
            SizeT | None: Size style or ``None`` if cancelled.
        """
        return self.__styler.style_get(factory=chart2_position_size_size_factory)
