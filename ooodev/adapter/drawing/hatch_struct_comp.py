from __future__ import annotations
from typing import TYPE_CHECKING
import uno
from ooo.dyn.drawing.hatch import Hatch
from ooo.dyn.drawing.hatch_style import HatchStyle

from ooodev.adapter.struct_base import StructBase
from ooodev.units.angle10 import Angle10

if TYPE_CHECKING:
    from ooodev.events.events_t import EventsT
    from ooodev.utils.color import Color
    from ooodev.units.angle_t import AngleT


class HatchStructComp(StructBase[Hatch]):
    """
    Hatch Struct

    This class raises an event before and after a property is changed if it has been passed an event provider.

    The event raised before the property is changed is called ``com_sun_star_drawing_Hatch_changing``.
    The event raised after the property is changed is called ``com_sun_star_drawing_Hatch_changed``.

    The event args for before the property is changed is of type ``KeyValCancelArgs``.
    The event args for after the property is changed is of type ``KeyValArgs``.
    """

    def __init__(self, component: Hatch, prop_name: str, event_provider: EventsT | None = None) -> None:
        """
        Constructor

        Args:
            component (Hatch): Border Line.
            prop_name (str): Property Name. This value is assigned to the ``prop_name`` of ``event_data``.
            event_provider (EventsT, optional): Event Provider.
        """
        super().__init__(component=component, prop_name=prop_name, event_provider=event_provider)

    # region Overrides
    def _get_on_changing_event_name(self) -> str:
        return "com_sun_star_drawing_Hatch_changing"

    def _get_on_changed_event_name(self) -> str:
        return "com_sun_star_drawing_Hatch_changed"

    def _copy(self, src: Hatch | None = None) -> Hatch:
        if src is None:
            src = self.component
        return Hatch(
            Style=src.Style,
            Color=src.Color,
            Distance=src.Distance,
            Angle=src.Angle,
        )

    # endregion Overrides
    @property
    def style(self) -> HatchStyle:
        """
        Gets/Sets the style of the gradient.

        Returns:
            HatchStyle: Hatch Style

        Hint:
            - ``HatchStyle`` can be imported from ``ooodev.drawing.hatch_style``
        """
        return self.component.Style  # type: ignore

    @style.setter
    def style(self, value: HatchStyle) -> None:
        old_value = self.component.Style
        if old_value != value:
            event_args = self._trigger_cancel_event("Style", old_value, value)
            _ = self._trigger_done_event(event_args)

    # region Properties
    @property
    def angle(self) -> Angle10:
        """
        Gets/Sets angle of the gradient in ``1/10`` degree.

        When setting the value can be set with an ``int`` in ``1/10`` degrees or an ``AngleT`` instance.

        Returns:
            Angle10: Angle in ``1/10`` degree.

        Hint:
            - ``Angle10`` can be imported from ``ooodev.units``.
        """
        return Angle10(self.component.Angle)

    @angle.setter
    def angle(self, value: int | AngleT) -> None:
        old_value = self.component.Angle
        val = Angle10.from_unit_val(value)
        new_value = val.value
        if old_value != new_value:
            event_args = self._trigger_cancel_event("Angle", old_value, new_value)
            self._trigger_done_event(event_args)

    @property
    def color(self) -> Color:
        """
        Gets/Sets the color of the hatch lines.

        Returns:
            ~ooodev.utils.color.Color: Color
        """
        return self.component.Color  # type: ignore

    @color.setter
    def color(self, value: Color) -> None:
        old_value = self.component.Color
        if old_value != value:
            event_args = self._trigger_cancel_event("Color", old_value, value)
            _ = self._trigger_done_event(event_args)

    @property
    def distance(self) -> int:
        """
        This is the distance between the lines in the hatch.
        """
        return self.component.Distance

    @distance.setter
    def distance(self, value: int) -> None:
        old_value = self.component.Distance
        if old_value != value:
            event_args = self._trigger_cancel_event("Distance", old_value, value)
            _ = self._trigger_done_event(event_args)

    # endregion Properties
