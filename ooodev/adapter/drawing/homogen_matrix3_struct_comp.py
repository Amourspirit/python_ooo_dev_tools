from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from ooo.dyn.drawing.homogen_matrix3 import HomogenMatrix3
from ooo.dyn.drawing.homogen_matrix_line3 import HomogenMatrixLine3

from ooodev.adapter.drawing.homogen_matrix_line3_struct_comp import HomogenMatrixLine3StructComp
from ooodev.adapter.struct_base import StructBase
from ooodev.events.events import Events
from ooodev.events.args.key_val_args import KeyValArgs
from ooodev.utils import info as mInfo

if TYPE_CHECKING:
    from ooodev.events.events_t import EventsT


class HomogenMatrix3StructComp(StructBase[HomogenMatrix3]):
    """
    HomogenMatrix3 Struct

    This class raises an event before and after a property is changed if it has been passed an event provider.

    The event raised before the property is changed is called ``com_sun_star_drawing_HomogenMatrix3_changing``.
    The event raised after the property is changed is called ``com_sun_star_drawing_HomogenMatrix3_changed``.

    The event args for before the property is changed is of type ``KeyValCancelArgs``.
    The event args for after the property is changed is of type ``KeyValArgs``.
    """

    def __init__(self, component: HomogenMatrix3, prop_name: str, event_provider: EventsT | None) -> None:
        """
        Constructor

        Args:
            component (HomogenMatrix3): Homogen Matrix 3
            prop_name (str): Property Name. This value is assigned to the ``prop_name`` of ``event_data``.
            event_provider (EventsT | None): Event Provider.
        """
        super().__init__(component=component, prop_name=prop_name, event_provider=event_provider)
        self._props = {}
        self._events = Events(source=self)

        # pylint: disable=unused-argument

        def on_changed(src: Any, event_args: KeyValArgs) -> None:
            prop_name = str(event_args.event_data["prop_name"])
            if hasattr(self, prop_name):
                setattr(self, prop_name, event_args.source.component)

        self._fn_on_changed = on_changed
        self._events.subscribe_event("com_sun_star_drawing_HomogenMatrixLine3_changed", self._fn_on_changed)

    # region Overrides
    def _get_on_changing_event_name(self) -> str:
        return "com_sun_star_drawing_HomogenMatrix3_changing"

    def _get_on_changed_event_name(self) -> str:
        return "com_sun_star_drawing_HomogenMatrix3_changed"

    def _copy(self, src: HomogenMatrix3 | None = None) -> HomogenMatrix3:
        if src is None:
            src = self.component

        def copy_line(src: HomogenMatrixLine3) -> HomogenMatrixLine3:
            return HomogenMatrixLine3(
                Column1=src.Column1,
                Column2=src.Column2,
                Column3=src.Column3,
            )

        return HomogenMatrix3(
            Line1=copy_line(src.Line1),
            Line2=copy_line(src.Line2),
            Line3=copy_line(src.Line3),
        )

    # endregion Overrides

    # region Properties

    @property
    def line1(self) -> HomogenMatrixLine3StructComp:
        """
        Gets/Sets Line 1.

        Setting value can be done with a ``HomogenMatrixLine3`` or ``HomogenMatrixLine3StructComp`` object.

        Returns:
            HomogenMatrixLine3StructComp: Returns Homogen Matrix Line

        Hint:
            - ``HomogenMatrixLine3`` can be imported from `` ooo.dyn.drawing.homogen_matrix_line3``
        """
        key = "line1"
        prop = self._props.get(key, None)
        if prop is None:
            prop = HomogenMatrixLine3StructComp(self.component.Line1, key, self._event_provider)
            self._props[key] = prop
        return cast(HomogenMatrixLine3StructComp, prop)

    @line1.setter
    def line1(self, value: HomogenMatrixLine3StructComp | HomogenMatrixLine3) -> None:
        key = "line1"
        old_value = self.component.Line1
        if mInfo.Info.is_instance(value, HomogenMatrixLine3StructComp):
            new_value = value.copy()
        else:
            comp = HomogenMatrixLine3StructComp(cast(HomogenMatrixLine3, value), key)
            new_value = comp.copy()

        event_args = self._trigger_cancel_event("Line1", old_value, new_value)
        done_args = self._trigger_done_event(event_args)
        if done_args is None:
            return
        if key in self._props:
            del self._props[key]

    @property
    def line2(self) -> HomogenMatrixLine3StructComp:
        """
        Gets/Sets Line 2.

        Setting value can be done with a ``HomogenMatrixLine3`` or ``HomogenMatrixLine3StructComp`` object.

        Returns:
            HomogenMatrixLine3StructComp: Returns Homogen Matrix Line

        Hint:
            - ``HomogenMatrixLine3`` can be imported from `` ooo.dyn.drawing.homogen_matrix_line3``
        """
        key = "line2"
        prop = self._props.get(key, None)
        if prop is None:
            prop = HomogenMatrixLine3StructComp(self.component.Line2, key, self._event_provider)
            self._props[key] = prop
        return cast(HomogenMatrixLine3StructComp, prop)

    @line2.setter
    def line2(self, value: HomogenMatrixLine3StructComp | HomogenMatrixLine3) -> None:
        key = "line2"
        old_value = self.component.Line2
        if mInfo.Info.is_instance(value, HomogenMatrixLine3StructComp):
            new_value = value.copy()
        else:
            comp = HomogenMatrixLine3StructComp(cast(HomogenMatrixLine3, value), key)
            new_value = comp.copy()

        event_args = self._trigger_cancel_event("Line2", old_value, new_value)
        done_args = self._trigger_done_event(event_args)
        if done_args is None:
            return
        if key in self._props:
            del self._props[key]

    @property
    def line3(self) -> HomogenMatrixLine3StructComp:
        """
        Gets/Sets Line 3.

        Setting value can be done with a ``HomogenMatrixLine3`` or ``HomogenMatrixLine3StructComp`` object.

        Returns:
            HomogenMatrixLine3StructComp: Returns Homogen Matrix Line

        Hint:
            - ``HomogenMatrixLine3`` can be imported from `` ooo.dyn.drawing.homogen_matrix_line3``
        """
        key = "line3"
        prop = self._props.get(key, None)
        if prop is None:
            prop = HomogenMatrixLine3StructComp(self.component.Line3, key, self._event_provider)
            self._props[key] = prop
        return cast(HomogenMatrixLine3StructComp, prop)

    @line3.setter
    def line3(self, value: HomogenMatrixLine3StructComp | HomogenMatrixLine3) -> None:
        key = "line3"
        old_value = self.component.Line3
        if mInfo.Info.is_instance(value, HomogenMatrixLine3StructComp):
            new_value = value.copy()
        else:
            comp = HomogenMatrixLine3StructComp(cast(HomogenMatrixLine3, value), key)
            new_value = comp.copy()

        event_args = self._trigger_cancel_event("Line3", old_value, new_value)
        done_args = self._trigger_done_event(event_args)
        if done_args is None:
            return
        if key in self._props:
            del self._props[key]

    # endregion Properties
