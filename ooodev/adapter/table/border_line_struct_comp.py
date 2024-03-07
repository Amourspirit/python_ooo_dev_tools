from __future__ import annotations
from typing import TYPE_CHECKING
import uno
from ooo.dyn.table.border_line import BorderLine
from ooodev.adapter.struct_base import StructBase
from ooodev.units.unit_mm100 import UnitMM100

if TYPE_CHECKING:
    from ooodev.events.events_t import EventsT
    from ooodev.utils.color import Color
    from ooodev.units.unit_obj import UnitT

# It seems that it is necessary to assign the struct to a variable, then change the variable and assign it back to the component.
# It is as if LibreOffice creates a new instance of the struct when it is changed.


class BorderLineStructComp(StructBase[BorderLine]):
    """
    Border Line Struct

    This class raises an event before and after a property is changed if it has been passed an event provider.

    The event raised before the property is changed is called ``com_sun_star_table_BorderLine_changing``.
    The event raised after the property is changed is called ``com_sun_star_table_BorderLine_changed``.

    The event args for before the property is changed is of type ``KeyValCancelArgs``.
    The event args for after the property is changed is of type ``KeyValArgs``.
    """

    def __init__(self, component: BorderLine, prop_name: str, event_provider: EventsT | None = None) -> None:
        """
        Constructor

        Args:
            component (BorderLine): Border Line.
            prop_name (str): Property Name. This value is assigned to the ``prop_name`` of ``event_data``.
            event_provider (EventsT, optional): Event Provider.
        """
        super().__init__(component=component, prop_name=prop_name, event_provider=event_provider)

    # region Overrides
    def _get_on_changing_event_name(self) -> str:
        return "com_sun_star_table_BorderLine_changing"

    def _get_on_changed_event_name(self) -> str:
        return "com_sun_star_table_BorderLine_changed"

    def _copy(self, src: BorderLine | None = None) -> BorderLine:
        if src is None:
            src = self.component
        return BorderLine(
            Color=src.Color,
            InnerLineWidth=src.InnerLineWidth,
            OuterLineWidth=src.OuterLineWidth,
            LineDistance=src.LineDistance,
        )

    # endregion Overrides

    # region Properties

    @property
    def color(self) -> Color:
        """
        Gets/Sets the color value of the line.

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
    def inner_line_width(self) -> UnitMM100:
        """
        Gets/Sets the width of the inner part of a double line (in ``1/100 mm``).

        Value can be set with a ``UnitT`` instance or an ``int`` in ``1/100 mm`` units.

        Returns:
            UnitMM100: Inner Line Width.

        Hint:
            - ``UnitMM100`` can be imported from ``ooodev.units``.
        """
        return UnitMM100(self.component.InnerLineWidth)

    @inner_line_width.setter
    def inner_line_width(self, value: int | UnitT) -> None:
        old_value = self.component.InnerLineWidth
        val = UnitMM100.from_unit_val(value)
        new_value = val.value
        if old_value != new_value:
            event_args = self._trigger_cancel_event("InnerLineWidth", old_value, new_value)
            self._trigger_done_event(event_args)

    @property
    def outer_line_width(self) -> UnitMM100:
        """
        Gets/Sets the width of a single line or the width of outer part of a double line (in ``1/100 mm``).

        Value can be set with a ``UnitT`` instance or an ``int`` in ``1/100 mm`` units.

        Returns:
            UnitMM100: Outer Line Width.

        Hint:
            - ``UnitMM100`` can be imported from ``ooodev.units``.
        """
        return UnitMM100(self.component.OuterLineWidth)

    @outer_line_width.setter
    def outer_line_width(self, value: int | UnitT) -> None:
        old_value = self.component.OuterLineWidth
        val = UnitMM100.from_unit_val(value)
        new_value = val.value
        if old_value != new_value:
            event_args = self._trigger_cancel_event("OuterLineWidth", old_value, new_value)
            _ = self._trigger_done_event(event_args)

    @property
    def line_distance(self) -> UnitMM100:
        """
        Gets/Sets the distance between the inner and outer parts of a double line (in ``1/100 mm``).

        Value can be set with a ``UnitT`` instance or an ``int`` in ``1/100 mm`` units.

        Returns:
            UnitMM100: Line Distance.

        Hint:
            - ``UnitMM100`` can be imported from ``ooodev.units``.
        """
        return UnitMM100(self.component.LineDistance)

    @line_distance.setter
    def line_distance(self, value: int | UnitT) -> None:
        old_value = self.component.LineDistance
        val = UnitMM100.from_unit_val(value)
        new_value = val.value
        if old_value != new_value:
            event_args = self._trigger_cancel_event("LineDistance", old_value, new_value)
            _ = self._trigger_done_event(event_args)

    # endregion Properties
