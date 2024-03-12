from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from ooo.dyn.drawing.homogen_matrix import HomogenMatrix
from ooo.dyn.drawing.homogen_matrix_line import HomogenMatrixLine

from ooodev.adapter.drawing.homogen_matrix_line_struct_comp import HomogenMatrixLineStructComp
from ooodev.adapter.struct_base import StructBase
from ooodev.events.events import Events
from ooodev.events.args.key_val_args import KeyValArgs
from ooodev.utils import info as mInfo

if TYPE_CHECKING:
    from ooodev.events.events_t import EventsT


class HomogenMatrixStructComp(StructBase[HomogenMatrix]):
    """
    HomogenMatrix Struct

    This class raises an event before and after a property is changed if it has been passed an event provider.

    The event raised before the property is changed is called ``com_sun_star_drawing_HomogenMatrix_changing``.
    The event raised after the property is changed is called ``com_sun_star_drawing_HomogenMatrix_changed``.

    The event args for before the property is changed is of type ``KeyValCancelArgs``.
    The event args for after the property is changed is of type ``KeyValArgs``.
    """

    def __init__(self, component: HomogenMatrix, prop_name: str, event_provider: EventsT | None) -> None:
        """
        Constructor

        Args:
            component (HomogenMatrix): Homogen Matrix
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
        self._events.subscribe_event("com_sun_star_drawing_HomogenMatrixLine_changed", self._fn_on_changed)

    # region Overrides
    def _get_on_changing_event_name(self) -> str:
        return "com_sun_star_drawing_HomogenMatrix_changing"

    def _get_on_changed_event_name(self) -> str:
        return "com_sun_star_drawing_HomogenMatrix_changed"

    def _copy(self, src: HomogenMatrix | None = None) -> HomogenMatrix:
        if src is None:
            src = self.component

        def copy_line(src: HomogenMatrixLine) -> HomogenMatrixLine:
            return HomogenMatrixLine(
                Column1=src.Column1,
                Column2=src.Column2,
                Column3=src.Column3,
                Column4=src.Column4,
            )

        return HomogenMatrix(
            Line1=copy_line(src.Line1),
            Line2=copy_line(src.Line2),
            Line3=copy_line(src.Line3),
            Line4=copy_line(src.Line4),
        )

    # endregion Overrides

    # region Properties

    @property
    def line1(self) -> HomogenMatrixLineStructComp:
        """
        Gets/Sets Line 1.

        Setting value can be done with a ``HomogenMatrixLine`` or ``HomogenMatrixLineStructComp`` object.

        Returns:
            HomogenMatrixLineStructComp: Returns Homogen Matrix Line

        Hint:
            - ``HomogenMatrixLine`` can be imported from ``ooo.dyn.drawing.homogen_matrix_line``
        """
        key = "line1"
        prop = self._props.get(key, None)
        if prop is None:
            prop = HomogenMatrixLineStructComp(self.component.Line1, key, self._event_provider)
            self._props[key] = prop
        return cast(HomogenMatrixLineStructComp, prop)

    @line1.setter
    def line1(self, value: HomogenMatrixLineStructComp | HomogenMatrixLine) -> None:
        key = "line1"
        old_value = self.component.Line1
        if mInfo.Info.is_instance(value, HomogenMatrixLineStructComp):
            new_value = value.copy()
        else:
            comp = HomogenMatrixLineStructComp(cast(HomogenMatrixLine, value), key)
            new_value = comp.copy()

        event_args = self._trigger_cancel_event("Line1", old_value, new_value)
        done_args = self._trigger_done_event(event_args)
        if done_args is None:
            return
        if key in self._props:
            del self._props[key]

    @property
    def line2(self) -> HomogenMatrixLineStructComp:
        """
        Gets/Sets Line 2.

        Setting value can be done with a ``HomogenMatrixLine`` or ``HomogenMatrixLineStructComp`` object.

        Returns:
            HomogenMatrixLineStructComp: Returns Homogen Matrix Line

        Hint:
            - ``HomogenMatrixLine`` can be imported from ``ooo.dyn.drawing.homogen_matrix_line``
        """
        key = "line2"
        prop = self._props.get(key, None)
        if prop is None:
            prop = HomogenMatrixLineStructComp(self.component.Line2, key, self._event_provider)
            self._props[key] = prop
        return cast(HomogenMatrixLineStructComp, prop)

    @line2.setter
    def line2(self, value: HomogenMatrixLineStructComp | HomogenMatrixLine) -> None:
        key = "line2"
        old_value = self.component.Line2
        if mInfo.Info.is_instance(value, HomogenMatrixLineStructComp):
            new_value = value.copy()
        else:
            comp = HomogenMatrixLineStructComp(cast(HomogenMatrixLine, value), key)
            new_value = comp.copy()

        event_args = self._trigger_cancel_event("Line2", old_value, new_value)
        done_args = self._trigger_done_event(event_args)
        if done_args is None:
            return
        if key in self._props:
            del self._props[key]

    @property
    def line3(self) -> HomogenMatrixLineStructComp:
        """
        Gets/Sets Line 3.

        Setting value can be done with a ``HomogenMatrixLine`` or ``HomogenMatrixLineStructComp`` object.

        Returns:
            HomogenMatrixLineStructComp: Returns Homogen Matrix Line

        Hint:
            - ``HomogenMatrixLine`` can be imported from ``ooo.dyn.drawing.homogen_matrix_line``
        """
        key = "line3"
        prop = self._props.get(key, None)
        if prop is None:
            prop = HomogenMatrixLineStructComp(self.component.Line3, key, self._event_provider)
            self._props[key] = prop
        return cast(HomogenMatrixLineStructComp, prop)

    @line3.setter
    def line3(self, value: HomogenMatrixLineStructComp | HomogenMatrixLine) -> None:
        key = "line3"
        old_value = self.component.Line3
        if mInfo.Info.is_instance(value, HomogenMatrixLineStructComp):
            new_value = value.copy()
        else:
            comp = HomogenMatrixLineStructComp(cast(HomogenMatrixLine, value), key)
            new_value = comp.copy()

        event_args = self._trigger_cancel_event("Line3", old_value, new_value)
        done_args = self._trigger_done_event(event_args)
        if done_args is None:
            return
        if key in self._props:
            del self._props[key]

    @property
    def line4(self) -> HomogenMatrixLineStructComp:
        """
        Gets/Sets Line 4.

        Setting value can be done with a ``HomogenMatrixLine`` or ``HomogenMatrixLineStructComp`` object.

        Returns:
            HomogenMatrixLineStructComp: Returns Homogen Matrix Line

        Hint:
            - ``HomogenMatrixLine`` can be imported from ``ooo.dyn.drawing.homogen_matrix_line``
        """
        key = "line4"
        prop = self._props.get(key, None)
        if prop is None:
            prop = HomogenMatrixLineStructComp(self.component.Line4, key, self._event_provider)
            self._props[key] = prop
        return cast(HomogenMatrixLineStructComp, prop)

    @line4.setter
    def line4(self, value: HomogenMatrixLineStructComp | HomogenMatrixLine) -> None:
        key = "line4"
        old_value = self.component.Line4
        if mInfo.Info.is_instance(value, HomogenMatrixLineStructComp):
            new_value = value.copy()
        else:
            comp = HomogenMatrixLineStructComp(cast(HomogenMatrixLine, value), key)
            new_value = comp.copy()

        event_args = self._trigger_cancel_event("Line4", old_value, new_value)
        done_args = self._trigger_done_event(event_args)
        if done_args is None:
            return
        if key in self._props:
            del self._props[key]

    # endregion Properties
