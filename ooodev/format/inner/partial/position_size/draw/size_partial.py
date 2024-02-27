from __future__ import annotations
from typing import Any, TYPE_CHECKING

from ooodev.events.partial.events_partial import EventsPartial
from ooodev.format.inner.partial.factory_styler import FactoryStyler
from ooodev.format.inner.style_factory import draw_position_size_size_factory
from ooodev.loader import lo as mLo
from ooodev.utils.kind.shape_base_point_kind import ShapeBasePointKind

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.format.proto.draw.position_size.size_t import SizeT
    from ooodev.units.unit_obj import UnitT
else:
    SizeT = Any
    UnitT = Any


class SizePartial:
    """
    Partial class for Size.
    """

    def __init__(self, factory_name: str, component: Any, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self.__styler = FactoryStyler(factory_name=factory_name, component=component, lo_inst=lo_inst)
        if isinstance(self, EventsPartial):
            self.__styler.add_event_observers(self.event_observer)
        self.__styler.after_event_name = "after_style_size"
        self.__styler.before_event_name = "before_style_size"

    def style_size(
        self, width: float | UnitT, height: float | UnitT, base_point: ShapeBasePointKind = ShapeBasePointKind.TOP_LEFT
    ) -> SizeT | None:
        """
        Style Area Color.

        Args:
            width (float | UnitT): Specifies the width of the shape (in ``mm`` units) or :ref:`proto_unit_obj`.
            height (float | UnitT): Specifies the height of the shape (in ``mm`` units) or :ref:`proto_unit_obj`.
            base_point (ShapeBasePointKind, optional): Specifies the base point of the shape used to calculate the X and Y coordinates. Default is ``TOP_LEFT``.

        Raises:
            CancelEventError: If the event ``before_style_size`` is cancelled and not handled.

        Returns:
            SizeT | None: Position instance or ``None`` if cancelled.

        Hint:
            - ``ShapeBasePointKind`` can be imported from ``ooodev.utils.kind.shape_base_point_kind``
        """
        return self.__styler.style(
            factory=draw_position_size_size_factory, width=width, height=height, base_point=base_point
        )

    def style_size_get(self) -> SizeT | None:
        """
        Gets the Size Style.

        Raises:
            CancelEventError: If the event ``before_style_size_get`` is cancelled and not handled.

        Returns:
            SizeT | None: Size style or ``None`` if cancelled.
        """
        return self.__styler.style_get(factory=draw_position_size_size_factory)
