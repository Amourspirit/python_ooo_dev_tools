from __future__ import annotations
from typing import Any, TYPE_CHECKING

from ooodev.events.partial.events_partial import EventsPartial
from ooodev.format.inner.partial.factory_styler import FactoryStyler
from ooodev.format.inner.style_factory import draw_position_size_position_factory
from ooodev.loader import lo as mLo
from ooodev.utils.kind.shape_base_point_kind import ShapeBasePointKind

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.format.proto.draw.position_size.position_t import PositionT
    from ooodev.units import UnitT
else:
    PositionT = Any
    UnitT = Any


class PositionPartial:
    """
    Partial class for Position.
    """

    def __init__(self, factory_name: str, component: Any, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self.__styler = FactoryStyler(factory_name=factory_name, component=component, lo_inst=lo_inst)
        if isinstance(self, EventsPartial):
            self.__styler.add_event_observers(self.event_observer)
        self.__styler.after_event_name = "after_style_position"
        self.__styler.before_event_name = "before_style_position"

    def style_position(
        self, x: float | UnitT, y: float | UnitT, base_point: ShapeBasePointKind = ShapeBasePointKind.TOP_LEFT
    ) -> PositionT | None:
        """
        Style Area Color.

        Args:
            x (float, UnitT): Specifies the x-coordinate of the position of the shape (in ``mm`` units) or :ref:`proto_unit_obj`.
            y (float, UnitT): Specifies the y-coordinate of the position of the shape (in ``mm`` units) or :ref:`proto_unit_obj`.
            base_point (ShapeBasePointKind, optional): Specifies the base point of the shape used to calculate the X and Y coordinates. Default is ``TOP_LEFT``.

        Raises:
            CancelEventError: If the event ``before_style_position`` is cancelled and not handled.

        Returns:
            PositionT | None: Position instance or ``None`` if cancelled.

        Hint:
            - ``ShapeBasePointKind`` can be imported from ``ooodev.utils.kind.shape_base_point_kind``
        """
        return self.__styler.style(
            factory=draw_position_size_position_factory, pos_x=x, pos_y=y, base_point=base_point
        )

    def style_position_get(self) -> PositionT | None:
        """
        Gets the Position Style.

        Raises:
            CancelEventError: If the event ``before_style_position_get`` is cancelled and not handled.

        Returns:
            PositionT | None: Position style or ``None`` if cancelled.
        """
        return self.__styler.style_get(factory=draw_position_size_position_factory)
