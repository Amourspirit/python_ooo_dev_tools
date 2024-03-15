from __future__ import annotations
from typing import cast, Generic, Type, TypeVar, TYPE_CHECKING
import uno
from ooo.dyn.text.graphic_crop import GraphicCrop

from ooodev.units.unit_mm100 import UnitMM100
from ooodev.adapter.struct_base import StructBase
from ooodev.units.unit_factory import get_unit

if TYPE_CHECKING:
    from ooodev.events.events_t import EventsT
    from ooodev.units.unit_obj import UnitT

_T = TypeVar("_T", bound="UnitT")


class GraphicCropStructComp(StructBase[GraphicCrop], Generic[_T]):
    """
    Generic GraphicCrop Struct

    This class raises an event before and after a property is changed if it has been passed an event provider.

    The event raised before the property is changed is called ``com_sun_star_text_GraphicCrop_changing``.
    The event raised after the property is changed is called ``com_sun_star_text_GraphicCrop_changed``.

    The event args for before the property is changed is of type ``KeyValCancelArgs``.
    The event args for after the property is changed is of type ``KeyValArgs``.
    """

    def __init__(
        self, component: GraphicCrop, unit: Type[_T], prop_name: str, event_provider: EventsT | None = None
    ) -> None:
        """
        Constructor

        Args:
            component (GraphicCrop): Graphic Crop.
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
        return "com_sun_star_text_GraphicCrop_changing"

    def _get_on_changed_event_name(self) -> str:
        return "com_sun_star_text_GraphicCrop_changed"

    def _copy(self, src: GraphicCrop | None = None) -> GraphicCrop:
        if src is None:
            src = self.component
        return GraphicCrop(
            Top=src.Top,
            Bottom=src.Bottom,
            Left=src.Left,
            Right=src.Right,
        )

    # endregion Overrides
    # region Properties

    @property
    def top(self) -> _T:
        """
        Gets/Sets the top value to cut (if negative) or to extend (if positive)

        When setting the value can be a ``int`` in ``1/100th mm`` units or a ``UnitT`` measurement unit.

        Returns:
            _T: ``UnitT`` measurement unit.
        """
        unit100 = UnitMM100(self.component.Top)
        if not self._require_convert:
            return cast(_T, unit100)
        val = unit100.convert_to(self._unit_length)
        return cast(_T, get_unit(self._unit_length, val))

    @top.setter
    def top(self, value: _T | float) -> None:
        val = UnitMM100.from_unit_val(value)
        new_value = val.value
        old_value = self.component.Top
        if old_value != new_value:
            event_args = self._trigger_cancel_event("Top", old_value, new_value)
            self._trigger_done_event(event_args)

    @property
    def bottom(self) -> _T:
        """
        Gets/Sets the bottom value to cut (if negative) or to extend (if positive)

        When setting the value can be a ``int`` in ``1/100th mm`` units or a ``UnitT`` measurement unit.

        Returns:
            _T: ``UnitT`` measurement unit.
        """
        unit100 = UnitMM100(self.component.Bottom)
        if not self._require_convert:
            return cast(_T, unit100)
        val = unit100.convert_to(self._unit_length)
        return cast(_T, get_unit(self._unit_length, val))

    @bottom.setter
    def bottom(self, value: _T | float) -> None:
        val = UnitMM100.from_unit_val(value)
        new_value = val.value
        old_value = self.component.Bottom
        if old_value != new_value:
            event_args = self._trigger_cancel_event("Bottom", old_value, new_value)
            self._trigger_done_event(event_args)

    @property
    def left(self) -> _T:
        """
        Gets/Sets the left value to cut (if negative) or to extend (if positive)

        When setting the value can be a ``int`` in ``1/100th mm`` units or a ``UnitT`` measurement unit.

        Returns:
            _T: ``UnitT`` measurement unit.
        """
        unit100 = UnitMM100(self.component.Left)
        if not self._require_convert:
            return cast(_T, unit100)
        val = unit100.convert_to(self._unit_length)
        return cast(_T, get_unit(self._unit_length, val))

    @left.setter
    def left(self, value: _T | float) -> None:
        val = UnitMM100.from_unit_val(value)
        new_value = val.value
        old_value = self.component.Left
        if old_value != new_value:
            event_args = self._trigger_cancel_event("Left", old_value, new_value)
            self._trigger_done_event(event_args)

    @property
    def right(self) -> _T:
        """
        Gets/Sets the right value to cut (if negative) or to extend (if positive)

        When setting the value can be a ``int`` in ``1/100th mm`` units or a ``UnitT`` measurement unit.

        Returns:
            _T: ``UnitT`` measurement unit.
        """
        unit100 = UnitMM100(self.component.Right)
        if not self._require_convert:
            return cast(_T, unit100)
        val = unit100.convert_to(self._unit_length)
        return cast(_T, get_unit(self._unit_length, val))

    @right.setter
    def right(self, value: _T | float) -> None:
        val = UnitMM100.from_unit_val(value)
        new_value = val.value
        old_value = self.component.Right
        if old_value != new_value:
            event_args = self._trigger_cancel_event("Right", old_value, new_value)
            self._trigger_done_event(event_args)

    # endregion Properties
