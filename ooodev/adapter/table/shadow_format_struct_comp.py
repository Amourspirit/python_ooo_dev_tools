from __future__ import annotations
from typing import TYPE_CHECKING
import uno
from ooo.dyn.table.shadow_format import ShadowFormat
from ooo.dyn.table.shadow_location import ShadowLocation
from ooodev.adapter.struct_base import StructBase
from ooodev.units.unit_mm100 import UnitMM100

if TYPE_CHECKING:
    from ooodev.events.events_t import EventsT
    from ooodev.utils.color import Color
    from ooodev.units.unit_obj import UnitT

# It seems that it is necessary to assign the struct to a variable, then change the variable and assign it back to the component.
# It is as if LibreOffice creates a new instance of the struct when it is changed.


class ShadowFormatStructComp(StructBase[ShadowFormat]):
    """
    Shadow Format Struct.

    This class raises an event before and after a property is changed if it has been passed an event provider.

    The event raised before the property is changed is called ``com_sun_star_table_ShadowFormat_changing``.
    The event raised after the property is changed is called ``com_sun_star_table_ShadowFormat_changed``.

    The event args for before the property is changed is of type ``KeyValCancelArgs``.
    The event args for after the property is changed is of type ``KeyValArgs``.
    """

    def __init__(self, component: ShadowFormat, prop_name: str, event_provider: EventsT | None = None) -> None:
        """
        Constructor

        Args:
            component (ShadowFormat): Shadow Format.
            prop_name (str): Property Name. This value is assigned to the ``prop_name`` of ``event_data``.
            event_provider (EventsT, optional): Event Provider.
        """
        super().__init__(component=component, prop_name=prop_name, event_provider=event_provider)

    # region Overrides

    def _get_on_changed_event_name(self) -> str:
        return "com_sun_star_table_ShadowFormat_changed"

    def _get_on_changing_event_name(self) -> str:
        return "com_sun_star_table_ShadowFormat_changing"

    def _get_prop_name(self) -> str:
        return self._prop_name

    def _copy(self, src: ShadowFormat | None = None) -> ShadowFormat:
        if src is None:
            src = self.component
        return ShadowFormat(
            Location=src.Location,
            ShadowWidth=src.ShadowWidth,
            IsTransparent=src.IsTransparent,
            Color=src.Color,
        )

    # endregion Overrides

    # region Properties

    @property
    def location(self) -> ShadowLocation:
        """
        Gets/Sets the location of the shadow.

        Returns:
            ShadowLocation: Shadow Location.

        Hint:
            - ``ShadowLocation`` can be imported from ``ooo.dyn.table.shadow_location``.
        """
        return self.component.Location

    @location.setter
    def location(self, value: ShadowLocation) -> None:
        old_value = self.component.Location
        if old_value != value:
            event_args = self._trigger_cancel_event("Location", old_value, value)
            _ = self._trigger_done_event(event_args)

    @property
    def shadow_width(self) -> UnitMM100:
        """
        Gets/Sets the size of the shadow. (in ``1/100 mm``).

        Value can be set with a ``UnitT`` instance or an ``int`` in ``1/100 mm`` units.

        Returns:
            UnitMM100: Line Distance.

        Hint:
            - ``UnitMM100`` can be imported from ``ooodev.units.unit_mm100``.
        """
        return UnitMM100(self.component.ShadowWidth)

    @shadow_width.setter
    def shadow_width(self, value: int | UnitT) -> None:
        old_value = self.component.ShadowWidth
        val = UnitMM100.from_unit_val(value)
        new_value = val.value
        if old_value != new_value:
            event_args = self._trigger_cancel_event("ShadowWidth", old_value, new_value)
            _ = self._trigger_done_event(event_args)

    @property
    def is_transparent(self) -> bool:
        """
        Gets/Sets - ``True``, if shadow is transparent.

        Returns:
            bool: Is Transparent.
        """
        return self.component.IsTransparent  # type: ignore

    @is_transparent.setter
    def is_transparent(self, value: bool) -> None:
        old_value = self.component.IsTransparent
        if old_value != value:
            event_args = self._trigger_cancel_event("IsTransparent", old_value, value)
            _ = self._trigger_done_event(event_args)

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

    # endregion Properties
