from __future__ import annotations
from typing import TYPE_CHECKING
import uno
from ooo.dyn.drawing.homogen_matrix_line import HomogenMatrixLine

from ooodev.adapter.struct_base import StructBase

if TYPE_CHECKING:
    from ooodev.events.events_t import EventsT


class HomogenMatrixLineStructComp(StructBase[HomogenMatrixLine]):
    """
    HomogenMatrixLine Struct

    This class raises an event before and after a property is changed if it has been passed an event provider.

    The event raised before the property is changed is called ``com_sun_star_drawing_HomogenMatrixLine_changing``.
    The event raised after the property is changed is called ``com_sun_star_drawing_HomogenMatrixLine_changed``.

    The event args for before the property is changed is of type ``KeyValCancelArgs``.
    The event args for after the property is changed is of type ``KeyValArgs``.
    """

    def __init__(self, component: HomogenMatrixLine, prop_name: str, event_provider: EventsT | None = None) -> None:
        """
        Constructor

        Args:
            component (HomogenMatrixLine): Homogen Matrix Line.
            prop_name (str): Property Name. This value is assigned to the ``prop_name`` of ``event_data``.
            event_provider (EventsT, optional): Event Provider.
        """
        super().__init__(component=component, prop_name=prop_name, event_provider=event_provider)

    # region Overrides
    def _get_on_changing_event_name(self) -> str:
        return "com_sun_star_drawing_HomogenMatrixLine_changing"

    def _get_on_changed_event_name(self) -> str:
        return "com_sun_star_drawing_HomogenMatrixLine_changed"

    def _copy(self, src: HomogenMatrixLine | None = None) -> HomogenMatrixLine:
        if src is None:
            src = self.component
        return HomogenMatrixLine(
            Column1=src.Column1,
            Column2=src.Column2,
            Column3=src.Column3,
            Column4=src.Column4,
        )

    # endregion Overrides

    # region Properties
    @property
    def column1(self) -> float:
        """
        Gets/Sets the number of Column1 in this HomogenMatrixLine.
        """
        return self.component.Column1

    @column1.setter
    def column1(self, value: float) -> None:
        old_value = self.component.Column1
        if old_value != value:
            event_args = self._trigger_cancel_event("Column1", old_value, value)
            _ = self._trigger_done_event(event_args)

    @property
    def column2(self) -> float:
        """
        Gets/Sets the number of Column2 in this HomogenMatrixLine.
        """
        return self.component.Column2

    @column2.setter
    def column2(self, value: float) -> None:
        old_value = self.component.Column2
        if old_value != value:
            event_args = self._trigger_cancel_event("Column2", old_value, value)
            _ = self._trigger_done_event(event_args)

    @property
    def column3(self) -> float:
        """
        Gets/Sets the number of Column3 in this HomogenMatrixLine.
        """
        return self.component.Column3

    @column3.setter
    def column3(self, value: float) -> None:
        old_value = self.component.Column3
        if old_value != value:
            event_args = self._trigger_cancel_event("Column3", old_value, value)
            _ = self._trigger_done_event(event_args)

    @property
    def column4(self) -> float:
        """
        Gets/Sets the number of Column4 in this HomogenMatrixLine.
        """
        return self.component.Column4

    @column4.setter
    def column4(self, value: float) -> None:
        old_value = self.component.Column4
        if old_value != value:
            event_args = self._trigger_cancel_event("Column4", old_value, value)
            _ = self._trigger_done_event(event_args)

    # endregion Properties
