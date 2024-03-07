from __future__ import annotations
from typing import TYPE_CHECKING
import uno
from ooo.dyn.table.border_line2 import BorderLine2
from ooo.dyn.table.border_line_style import BorderLineStyleEnum
from ooodev.adapter.table.border_line_struct_comp import BorderLineStructComp
from ooodev.units.unit_mm100 import UnitMM100


if TYPE_CHECKING:
    from ooodev.events.events_t import EventsT
    from ooodev.units.unit_obj import UnitT

# It seems that it is necessary to assign the struct to a variable, then change the variable and assign it back to the component.
# It is as if LibreOffice creates a new instance of the struct when it is changed.


class BorderLine2StructComp(BorderLineStructComp):
    """
    Border2 Line Struct. A border line, extended with line style.

    This class raises an event before and after a property is changed if it has been passed an event provider.

    The event raised before the property is changed is called ``com_sun_star_table_BorderLine_changing``.
    The event raised after the property is changed is called ``com_sun_star_table_BorderLine_changed``.

    The event args for before the property is changed is of type ``KeyValCancelArgs``.
    The event args for after the property is changed is of type ``KeyValArgs``.
    """

    def __init__(self, component: BorderLine2, prop_name: str, event_provider: EventsT | None = None) -> None:
        """
        Constructor

        Args:
            component (BorderLine): Border Line 2.
            prop_name (str): Property Name. This value is assigned to the ``prop_name`` of ``event_data``.
            event_provider (EventsT, optional): Event Provider.
        """
        super().__init__(component, prop_name, event_provider)  # type: ignore

    # region Overrides

    def _get_on_changed_event_name(self) -> str:
        return "com_sun_star_table_BorderLine2_changed"

    def _get_on_changing_event_name(self) -> str:
        return "com_sun_star_table_BorderLine2_changing"

    def _copy(self, src: BorderLine2 | None = None) -> BorderLine2:
        if src is None:
            src = self.component
        return BorderLine2(
            Color=src.Color,
            InnerLineWidth=src.InnerLineWidth,
            OuterLineWidth=src.OuterLineWidth,
            LineDistance=src.LineDistance,
            LineWidth=src.LineWidth,
            LineStyle=src.LineStyle,
        )

    def copy(self) -> BorderLine2:
        """
        Makes a copy of the Border Line.

        Returns:
            BorderLine: Copied Border Line.
        """
        return self._copy()

    # endregion Overrides

    # region Properties

    @property
    def component(self) -> BorderLine2:
        """BorderLine Component"""
        # pylint: disable=no-member
        return self._get_component()  # type: ignore

    @component.setter
    def component(self, value: BorderLine2) -> None:
        # pylint: disable=no-member
        self._set_component(value, True)

    @property
    def line_style(self) -> BorderLineStyleEnum:
        """
        Gets/Sets the style of the border.

        When setting the value, it can be set with an ``int`` or a ``BorderLineStyleEnum`` instance.

        Returns:
            BorderLineStyleEnum: Border Line Style Enum.

        Hint:
            - ``BorderLineStyleEnum`` can be imported from ``ooo.dyn.table.border_line_style``.
        """
        return BorderLineStyleEnum(self.component.LineStyle)  # type: ignore

    @line_style.setter
    def line_style(self, value: int | BorderLineStyleEnum) -> None:
        old_value = self.component.LineStyle
        val = BorderLineStyleEnum(value)
        new_value = val.value
        if old_value != new_value:
            event_args = self._trigger_cancel_event("LineStyle", old_value, new_value)
            self._trigger_done_event(event_args)

    @property
    def line_width(self) -> UnitMM100:
        """
        Gets/Sets width of the border, this is the base to compute all the lines and gaps widths.
        These widths computations are based on the ``line_style`` property.
        This property is prevailing on the old Out, In and Dist width from ``border_line``.
        If this property is set to ``0``, then the other widths will be used to guess the border width.


        Value can be set with a ``UnitT`` instance or an ``int`` in ``1/100 mm`` units.

        Returns:
            UnitMM100: Line Width.

        Hint:
            - ``UnitMM100`` can be imported from ``ooodev.units.unit_mm100``.
            - ``UnitPT`` can be imported from ``ooodev.units.unit_pt``
            - ``HAIRLINE`` is ``UnitPT(0.05)``
            - ``VERY_THIN`` is ``UnitPT(0.5)``
            - ``THIN`` is ``UnitPT(0.75)``
            - ``MEDIUM`` is ``UnitPT(1.5)``
            - ``THICK`` is ``UnitPT(2.25)``
            - ``EXTRA_THICK`` is ``UnitPT(4.5)``
        """
        return UnitMM100(self.component.LineWidth)

    @line_width.setter
    def line_width(self, value: int | UnitT) -> None:
        old_value = self.component.LineWidth
        val = UnitMM100.from_unit_val(value)
        new_value = val.value
        if old_value != new_value:
            event_args = self._trigger_cancel_event("LineWidth", old_value, new_value)
            self._trigger_done_event(event_args)

    # endregion Properties
