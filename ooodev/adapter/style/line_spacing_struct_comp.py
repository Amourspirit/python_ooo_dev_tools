from __future__ import annotations
from typing import TYPE_CHECKING
import uno
from ooo.dyn.style.line_spacing import LineSpacing
from ooodev.adapter.struct_base import StructBase
from ooodev.units.unit_mm100 import UnitMM100
from ooodev.utils.kind.line_spacing_mode_kind import ModeKind

if TYPE_CHECKING:
    from ooodev.events.events_t import EventsT
    from ooodev.units.unit_obj import UnitT


class LineSpacingStructComp(StructBase[LineSpacing]):
    """
    Line Spacing Struct.

    This class raises an event before and after a property is changed if it has been passed an event provider.

    The event raised before the property is changed is called ``com_sun_star_style_LineSpacing_changing``.
    The event raised after the property is changed is called ``com_sun_star_style_LineSpacing_changed``.

    The event args for before the property is changed is of type ``KeyValCancelArgs``.
    The event args for after the property is changed is of type ``KeyValArgs``.
    """

    def __init__(self, component: LineSpacing, prop_name: str, event_provider: EventsT | None = None) -> None:
        """
        Constructor

        Args:
            component (LineSpacing): Line Spacing Component.
            prop_name (str): Property Name. This value is assigned to the ``prop_name`` of ``event_data``.
            event_provider (EventsT, optional): Event Provider.
        """
        super().__init__(component=component, prop_name=prop_name, event_provider=event_provider)

    # region Overrides
    def _get_on_changing_event_name(self) -> str:
        return "com_sun_star_style_LineSpacing_changing"

    def _get_on_changed_event_name(self) -> str:
        return "com_sun_star_style_LineSpacing_changed"

    def _copy(self, src: LineSpacing | None = None) -> LineSpacing:
        if src is None:
            src = self.component
        return LineSpacing(
            Mode=src.Mode,
            Height=src.Height,
        )

    # endregion Overrides

    # region Properties

    @property
    def mode(self) -> ModeKind:
        """
        Gets/Sets - This value specifies the way the height is specified.

        Returns:
            ModeKind: Mode Kind.

        Hint:
            - ``ModeKind`` can be imported from ``ooodev.utils.kind.line_spacing_mode_kind``
        """
        return ModeKind.from_uno(self.component)

    @mode.setter
    def mode(self, value: ModeKind) -> None:
        old_value = self.component.Mode
        new_value = value.get_mode()
        if old_value != new_value:
            event_args = self._trigger_cancel_event("Mode", old_value, new_value)
            self._trigger_done_event(event_args)

    @property
    def height(self) -> UnitMM100:
        """
        Gets/Sets the height in regard to Mode.

        When setting value can be an ``int`` in ``1/100th mm`` units or a ``UnitT``.

        Returns:
            UnitMM100: Height.

        Hint:
            - ``UnitMM100`` can be imported from ``ooodev.units``
        """
        return UnitMM100(self.component.Height)

    @height.setter
    def height(self, value: int | UnitT) -> None:
        old_value = self.component.Height
        val = UnitMM100.from_unit_val(value)
        new_value = val.value
        if old_value != new_value:
            event_args = self._trigger_cancel_event("Height", old_value, new_value)
            self._trigger_done_event(event_args)

    # endregion Properties
