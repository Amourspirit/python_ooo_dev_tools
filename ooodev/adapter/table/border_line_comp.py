from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from ooo.dyn.table.border_line import BorderLine
from ooodev.adapter.component_base import ComponentBase
from ooodev.events.args.key_val_cancel_args import KeyValCancelArgs
from ooodev.events.args.key_val_args import KeyValArgs
from ooodev.units.unit_mm100 import UnitMM100

if TYPE_CHECKING:
    from ooodev.events.events_t import EventsT
    from ooodev.utils.color import Color
    from ooodev.units.unit_obj import UnitT

# It seems that it is necessary to assign the struct to a variable, then change the variable and assign it back to the component.
# It is as if LibreOffice creates a new instance of the struct when it is changed.


class BorderLineComp(ComponentBase):
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
        ComponentBase.__init__(self, component)
        self._event_provider = event_provider
        self._prop_name = prop_name

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        # PropertySetPartial will validate
        return ()

    # endregion Overrides

    def _get_on_changed_event_name(self) -> str:
        return "com_sun_star_table_BorderLine_changed"

    def _get_on_changing_event_name(self) -> str:
        return "com_sun_star_table_BorderLine_changing"

    def _get_prop_name(self) -> str:
        return self._prop_name

    def _on_property_changing(self, event_args: KeyValCancelArgs) -> None:
        if self._event_provider is not None:
            self._event_provider.trigger_event(self._get_on_changing_event_name(), event_args)

    def _on_property_changed(self, event_args: KeyValArgs) -> None:
        if self._event_provider is not None:
            self._event_provider.trigger_event(self._get_on_changed_event_name(), event_args)

    def _copy(self, src: BorderLine | None = None) -> BorderLine:
        if src is None:
            src = self.component
        return BorderLine(
            Color=src.Color,
            InnerLineWidth=src.InnerLineWidth,
            OuterLineWidth=src.OuterLineWidth,
            LineDistance=src.LineDistance,
        )

    def copy(self) -> BorderLine:
        """
        Makes a copy of the Border Line.

        Returns:
            BorderLine: Copied Border Line.
        """
        return self._copy()

    # region Properties

    @property
    def component(self) -> BorderLine:
        """BorderLine Component"""
        # pylint: disable=no-member
        return cast("BorderLine", self._ComponentBase__get_component())  # type: ignore

    @component.setter
    def component(self, value: BorderLine) -> None:
        # pylint: disable=no-member
        self._ComponentBase__set_component(self._copy(src=value))  # type: ignore

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
            event_args = KeyValCancelArgs(
                source=self,
                key="color",
                value=value,
            )
            event_args.event_data = {
                "old_value": old_value,
                "prop_name": self._get_prop_name(),
            }
            self._on_property_changing(event_args)
            if not event_args.cancel:
                struct = self._copy()
                struct.Color = event_args.value
                self.component = struct
                self._on_property_changed(KeyValArgs.from_args(event_args))  # type: ignore

    @property
    def inner_line_width(self) -> UnitMM100:
        """
        Gets/Sets the width of the inner part of a double line (in ``1/100 mm``).

        Value can be set with a ``UnitT`` instance or an ``int`` in ``1/100 mm`` units.

        Returns:
            UnitMM100: Inner Line Width.

        Hint:
            - ``UnitMM100`` can be imported from ``ooodev.units.unit_mm100``.
        """
        return UnitMM100(self.component.InnerLineWidth)

    @inner_line_width.setter
    def inner_line_width(self, value: int | UnitT) -> None:
        old_value = self.component.InnerLineWidth
        val = UnitMM100.from_unit_val(value)
        new_value = val.value
        if old_value != new_value:
            event_args = KeyValCancelArgs(
                source=self,
                key="inner_line_width",
                value=new_value,
            )
            event_args.event_data = {
                "old_value": old_value,
                "prop_name": self._get_prop_name(),
            }
            self._on_property_changing(event_args)
            if not event_args.cancel:
                struct = self._copy()
                struct.InnerLineWidth = event_args.value
                self.component = struct
                self._on_property_changed(KeyValArgs.from_args(event_args))  # type: ignore

    @property
    def outer_line_width(self) -> UnitMM100:
        """
        Gets/Sets the width of a single line or the width of outer part of a double line (in ``1/100 mm``).

        Value can be set with a ``UnitT`` instance or an ``int`` in ``1/100 mm`` units.

        Returns:
            UnitMM100: Outer Line Width.

        Hint:
            - ``UnitMM100`` can be imported from ``ooodev.units.unit_mm100``.
        """
        return UnitMM100(self.component.OuterLineWidth)

    @outer_line_width.setter
    def outer_line_width(self, value: int | UnitT) -> None:
        old_value = self.component.OuterLineWidth
        val = UnitMM100.from_unit_val(value)
        new_value = val.value
        if old_value != new_value:
            event_args = KeyValCancelArgs(
                source=self,
                key="outer_line_width",
                value=new_value,
            )
            event_args.event_data = {
                "old_value": old_value,
                "prop_name": self._get_prop_name(),
            }
            self._on_property_changing(event_args)
            if not event_args.cancel:
                struct = self._copy()
                struct.OuterLineWidth = event_args.value
                self.component = struct
                self._on_property_changed(KeyValArgs.from_args(event_args))  # type: ignore

    @property
    def line_distance(self) -> UnitMM100:
        """
        Gets/Sets the distance between the inner and outer parts of a double line (in ``1/100 mm``).

        Value can be set with a ``UnitT`` instance or an ``int`` in ``1/100 mm`` units.

        Returns:
            UnitMM100: Line Distance.

        Hint:
            - ``UnitMM100`` can be imported from ``ooodev.units.unit_mm100``.
        """
        return UnitMM100(self.component.LineDistance)

    @line_distance.setter
    def line_distance(self, value: int | UnitT) -> None:
        old_value = self.component.LineDistance
        val = UnitMM100.from_unit_val(value)
        new_value = val.value
        if old_value != new_value:
            event_args = KeyValCancelArgs(
                source=self,
                key="line_distance",
                value=new_value,
            )
            event_args.event_data = {
                "old_value": old_value,
                "prop_name": self._get_prop_name(),
            }
            self._on_property_changing(event_args)
            if not event_args.cancel:
                struct = self._copy()
                struct.LineDistance = event_args.value
                self.component = struct
                self._on_property_changed(KeyValArgs.from_args(event_args))  # type: ignore

    # endregion Properties
