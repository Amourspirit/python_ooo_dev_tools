from __future__ import annotations
from typing import cast, Generic, Type, TypeVar, TYPE_CHECKING
import uno
from ooo.dyn.awt.size import Size
from ooodev.adapter.struct_base import StructBase
from ooodev.units.unit_mm100 import UnitMM100
from ooodev.units.unit_factory import get_unit

if TYPE_CHECKING:
    from ooodev.events.events_t import EventsT
    from ooodev.units.unit_obj import UnitT

_T = TypeVar("_T", bound="UnitT")


class SizeStructGenericComp(StructBase[Size], Generic[_T]):
    """
    Generic Size Struct.

    This class raises an event before and after a property is changed if it has been passed an event provider.

    The event raised before the property is changed is called ``generic_com_sun_star_awt_Size_changing``.
    The event raised after the property is changed is called ``generic_com_sun_star_awt_Size_changed``.

    The event args for before the property is changed is of type ``KeyValCancelArgs``.
    The event args for after the property is changed is of type ``KeyValArgs``.
    """

    def __init__(self, component: Size, unit: Type[_T], prop_name: str, event_provider: EventsT | None = None) -> None:
        """
        Constructor

        Args:
            component (Size): Size.
            unit (Type[UnitT]): Unit Type.
            prop_name (str): Property Name. This value is assigned to the ``prop_name`` of ``event_data``.
            event_provider (EventsT, optional): Event Provider.
        """
        super().__init__(component=component, prop_name=prop_name, event_provider=event_provider)
        self._unit = unit
        self._unit_length = unit.get_unit_length()
        self._require_convert = self._unit_length != UnitMM100.get_unit_length()

    def __repr__(self) -> str:
        return f"{self.__class__.__name__}[{self._unit.__name__}] {repr(self.component)}"

    # region Overrides
    def _get_on_changing_event_name(self) -> str:
        return "generic_com_sun_star_awt_Size_changing"

    def _get_on_changed_event_name(self) -> str:
        return "generic_com_sun_star_awt_Size_changed"

    def _copy(self, src: Size | None = None) -> Size:
        if src is None:
            src = self.component
        return Size(
            Width=src.Width,
            Height=src.Height,
        )

    # endregion Overrides

    # region Properties

    @property
    def width(self) -> _T:
        """
        Gets/Sets the Width.

        When setting the value can be a ``int`` in ``1/100th mm`` units or a ``UnitT`` measurement unit.

        Returns:
            _T: ``UnitT`` measurement unit.
        """
        unit100 = UnitMM100(self.component.Width)
        if not self._require_convert:
            return cast(_T, unit100)
        val = unit100.convert_to(self._unit_length)
        return cast(_T, get_unit(self._unit_length, val))

    @width.setter
    def width(self, value: _T | float) -> None:
        val = UnitMM100.from_unit_val(value)
        new_value = val.value
        old_value = self.component.Width
        if old_value != new_value:
            event_args = self._trigger_cancel_event("Width", old_value, new_value)
            self._trigger_done_event(event_args)

    @property
    def height(self) -> _T:
        """
        Gets/Sets the the Height.

        When setting the value can be a ``int`` in ``1/100th mm`` units or a ``UnitT`` measurement unit.

        Returns:
            _T: ``UnitT`` measurement unit.
        """
        unit100 = UnitMM100(self.component.Height)
        if not self._require_convert:
            return cast(_T, unit100)
        val = unit100.convert_to(self._unit_length)
        return cast(_T, get_unit(self._unit_length, val))

    @height.setter
    def height(self, value: _T | float) -> None:
        val = UnitMM100.from_unit_val(value)
        new_value = val.value
        old_value = self.component.Height
        if old_value != new_value:
            event_args = self._trigger_cancel_event("Height", old_value, new_value)
            self._trigger_done_event(event_args)

    # endregion Properties
