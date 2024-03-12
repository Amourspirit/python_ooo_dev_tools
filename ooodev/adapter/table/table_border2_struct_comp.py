from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from ooo.dyn.table.table_border2 import TableBorder2
from ooo.dyn.table.border_line2 import BorderLine2

from ooodev.adapter.struct_base import StructBase
from ooodev.events.events import Events
from ooodev.events.args.key_val_args import KeyValArgs
from ooodev.adapter.table.border_line2_struct_comp import BorderLine2StructComp
from ooodev.utils import info as mInfo
from ooodev.units.unit_mm100 import UnitMM100

if TYPE_CHECKING:
    from ooodev.events.events_t import EventsT
    from ooodev.units.unit_obj import UnitT

# It seems that it is necessary to assign the struct to a variable, then change the variable and assign it back to the component.
# It is as if LibreOffice creates a new instance of the struct when it is changed.


class TableBorder2StructComp(StructBase[TableBorder2]):
    """
    Table Border2 Struct

    This class raises an event before and after a property is changed if it has been passed an event provider.

    The event raised before the property is changed is called ``com_sun_star_table_TableBorder_changing``.
    The event raised after the property is changed is called ``com_sun_star_table_TableBorder_changed``.

    The event args for before the property is changed is of type ``KeyValCancelArgs``.
    The event args for after the property is changed is of type ``KeyValArgs``.
    """

    def __init__(self, component: TableBorder2, prop_name: str, event_provider: EventsT | None) -> None:
        """
        Constructor

        Args:
            component (TableBorder2): Table Border 2.
            prop_name (str): Property Name. This value is assigned to the ``prop_name`` of ``event_data``.
            event_provider (EventsT | None): Event Provider.
        """
        super().__init__(component=component, prop_name=prop_name, event_provider=event_provider)
        self._props = {}
        self._events = Events(source=self)

        # pylint: disable=unused-argument

        def on_border_line_changed(src: Any, event_args: KeyValArgs) -> None:
            prop_name = str(event_args.event_data["prop_name"])
            if hasattr(self, prop_name):
                setattr(self, prop_name, event_args.source.component)

        self.__fn_on_border_line_changed = on_border_line_changed
        self._events.subscribe_event("com_sun_star_table_BorderLine2_changed", self.__fn_on_border_line_changed)

    # region Overrides

    def _get_on_changing_event_name(self) -> str:
        return "com_sun_star_table_TableBorder2_changing"

    def _get_on_changed_event_name(self) -> str:
        return "com_sun_star_table_TableBorder2_changed"

    def _copy(self, src: TableBorder2 | None = None) -> TableBorder2:
        if src is None:
            src = self.component

        def copy_bdr(src: BorderLine2) -> BorderLine2:
            return BorderLine2(
                Color=src.Color,
                InnerLineWidth=src.InnerLineWidth,
                OuterLineWidth=src.OuterLineWidth,
                LineDistance=src.LineDistance,
                LineStyle=src.LineStyle,
                LineWidth=src.LineWidth,
            )

        return TableBorder2(
            TopLine=copy_bdr(src.TopLine),
            IsTopLineValid=src.IsTopLineValid,
            BottomLine=copy_bdr(src.BottomLine),
            IsBottomLineValid=src.IsBottomLineValid,
            LeftLine=copy_bdr(src.LeftLine),
            IsLeftLineValid=src.IsLeftLineValid,
            RightLine=copy_bdr(src.RightLine),
            IsRightLineValid=src.IsRightLineValid,
            HorizontalLine=copy_bdr(src.HorizontalLine),
            IsHorizontalLineValid=src.IsHorizontalLineValid,
            VerticalLine=copy_bdr(src.VerticalLine),
            IsVerticalLineValid=src.IsVerticalLineValid,
            Distance=src.Distance,
            IsDistanceValid=src.IsDistanceValid,
        )

    # endregion Overrides

    # region Properties

    @property
    def top_line(self) -> BorderLine2StructComp:
        """
        Gets/Set the line style at the top edge.

        Setting value can be done with a ``BorderLine2`` or ``BorderLine2Comp`` object.

        Returns:
            BorderLine2Comp: Returns Border Line.

        Hint:
            - ``BorderLine2`` can be imported from ``ooo.dyn.table.border_line2``
        """
        key = "top_line"
        prop = self._props.get(key, None)
        if prop is None:
            prop = BorderLine2StructComp(self.component.TopLine, key, self._event_provider)
            self._props[key] = prop
        return cast(BorderLine2StructComp, prop)

    @top_line.setter
    def top_line(self, value: BorderLine2StructComp | BorderLine2) -> None:
        key = "top_line"
        old_value = self.component.TopLine
        if mInfo.Info.is_instance(value, BorderLine2StructComp):
            new_value = value.copy()
        else:
            comp = BorderLine2StructComp(cast(BorderLine2, value), key)
            new_value = comp.copy()

        event_args = self._trigger_cancel_event("TopLine", old_value, new_value)
        done_args = self._trigger_done_event(event_args)
        if done_args is None:
            return
        if key in self._props:
            del self._props[key]

    @property
    def is_top_line_valid(self) -> bool:
        """
        Gets/Sets whether the value of ``TableBorder2.TopLine`` is used.
        """
        return self.component.IsTopLineValid

    @is_top_line_valid.setter
    def is_top_line_valid(self, value: bool) -> None:
        old_value = self.component.IsTopLineValid
        if old_value != value:
            event_args = self._trigger_cancel_event("IsTopLineValid", old_value, value)
            _ = self._trigger_done_event(event_args)

    @property
    def bottom_line(self) -> BorderLine2StructComp:
        """
        Gets/Set the line style at the bottom edge.

        Setting value can be done with a ``BorderLine2`` or ``BorderLine2Comp`` object.

        Returns:
            BorderLine2Comp: Returns Border Line.

        Hint:
            - ``BorderLine2`` can be imported from ``ooo.dyn.table.border_line2``
        """
        key = "bottom_line"
        prop = self._props.get(key, None)
        if prop is None:
            prop = BorderLine2StructComp(self.component.BottomLine, key, self._events)
            self._props[key] = prop
        return cast(BorderLine2StructComp, prop)

    @bottom_line.setter
    def bottom_line(self, value: BorderLine2StructComp | BorderLine2) -> None:
        key = "bottom_line"
        old_value = self.component.BottomLine
        if mInfo.Info.is_instance(value, BorderLine2StructComp):
            new_value = value.copy()
        else:
            comp = BorderLine2StructComp(cast(BorderLine2, value), key)
            new_value = comp.copy()

        event_args = self._trigger_cancel_event("BottomLine", old_value, new_value)
        done_args = self._trigger_done_event(event_args)
        if done_args is None:
            return
        if key in self._props:
            del self._props[key]

    @property
    def is_bottom_line_valid(self) -> bool:
        """
        Gets/Sets whether the value of ``TableBorder2.BottomLine`` is used.
        """
        return self.component.IsBottomLineValid

    @is_bottom_line_valid.setter
    def is_bottom_line_valid(self, value: bool) -> None:
        old_value = self.component.IsBottomLineValid
        if old_value != value:
            event_args = self._trigger_cancel_event("IsBottomLineValid", old_value, value)
            _ = self._trigger_done_event(event_args)

    @property
    def left_line(self) -> BorderLine2StructComp:
        """
        Gets/Set the line style at the left edge.

        Setting value can be done with a ``BorderLine2`` or ``BorderLine2Comp`` object.

        Returns:
            BorderLine2Comp: Returns Border Line.

        Hint:
            - ``BorderLine2`` can be imported from ``ooo.dyn.table.border_line2``
        """
        key = "left_line"
        prop = self._props.get(key, None)
        if prop is None:
            prop = BorderLine2StructComp(self.component.LeftLine, key, self._events)
            self._props[key] = prop
        return cast(BorderLine2StructComp, prop)

    @left_line.setter
    def left_line(self, value: BorderLine2StructComp | BorderLine2) -> None:
        key = "left_line"
        old_value = self.component.LeftLine
        if mInfo.Info.is_instance(value, BorderLine2StructComp):
            new_value = value.copy()
        else:
            comp = BorderLine2StructComp(cast(BorderLine2, value), key)
            new_value = comp.copy()

        event_args = self._trigger_cancel_event("LeftLine", old_value, new_value)
        done_args = self._trigger_done_event(event_args)
        if done_args is None:
            return
        if key in self._props:
            del self._props[key]

    @property
    def is_left_line_valid(self) -> bool:
        """
        Gets/Sets whether the value of ``TableBorder2.LeftLine`` is used.
        """
        return self.component.IsLeftLineValid

    @is_left_line_valid.setter
    def is_left_line_valid(self, value: bool) -> None:
        old_value = self.component.IsLeftLineValid
        if old_value != value:
            event_args = self._trigger_cancel_event("IsLeftLineValid", old_value, value)
            _ = self._trigger_done_event(event_args)

    @property
    def right_line(self) -> BorderLine2StructComp:
        """
        Gets/Set the line style at the right edge.

        Setting value can be done with a ``BorderLine2`` or ``BorderLine2Comp`` object.

        Returns:
            BorderLine2Comp: Returns Border Line.

        Hint:
            - ``BorderLine2`` can be imported from ``ooo.dyn.table.border_line2``
        """
        key = "right_line"
        prop = self._props.get(key, None)
        if prop is None:
            prop = BorderLine2StructComp(self.component.RightLine, key, self._events)
            self._props[key] = prop
        return cast(BorderLine2StructComp, prop)

    @right_line.setter
    def right_line(self, value: BorderLine2StructComp | BorderLine2) -> None:
        key = "right_line"
        old_value = self.component.RightLine
        if mInfo.Info.is_instance(value, BorderLine2StructComp):
            new_value = value.copy()
        else:
            comp = BorderLine2StructComp(cast(BorderLine2, value), key)
            new_value = comp.copy()

        event_args = self._trigger_cancel_event("RightLine", old_value, new_value)
        done_args = self._trigger_done_event(event_args)
        if done_args is None:
            return
        if key in self._props:
            del self._props[key]

    @property
    def is_right_line_valid(self) -> bool:
        """
        Gets/Sets whether the value of ``TableBorder2.RightLine`` is used.
        """
        return self.component.IsRightLineValid

    @is_right_line_valid.setter
    def is_right_line_valid(self, value: bool) -> None:
        old_value = self.component.IsRightLineValid
        if old_value != value:
            event_args = self._trigger_cancel_event("IsRightLineValid", old_value, value)
            _ = self._trigger_done_event(event_args)

    @property
    def horizontal_line(self) -> BorderLine2StructComp:
        """
        Gets/Set the horizontal line style edge.

        Setting value can be done with a ``BorderLine2`` or ``BorderLine2Comp`` object.

        Returns:
            BorderLine2Comp: Returns Border Line.

        Hint:
            - ``BorderLine2`` can be imported from ``ooo.dyn.table.border_line2``
        """
        key = "horizontal_line"
        prop = self._props.get(key, None)
        if prop is None:
            prop = BorderLine2StructComp(self.component.HorizontalLine, key, self._events)
            self._props[key] = prop
        return cast(BorderLine2StructComp, prop)

    @horizontal_line.setter
    def horizontal_line(self, value: BorderLine2StructComp | BorderLine2) -> None:
        key = "horizontal_line"
        old_value = self.component.HorizontalLine
        if mInfo.Info.is_instance(value, BorderLine2StructComp):
            new_value = value.copy()
        else:
            comp = BorderLine2StructComp(cast(BorderLine2, value), key)
            new_value = comp.copy()

        event_args = self._trigger_cancel_event("HorizontalLine", old_value, new_value)
        done_args = self._trigger_done_event(event_args)
        if done_args is None:
            return
        if key in self._props:
            del self._props[key]

    @property
    def is_horizontal_line_valid(self) -> bool:
        """
        Gets/Sets whether the value of ``TableBorder2.HorizontalLine`` is used.
        """
        return self.component.IsHorizontalLineValid

    @is_horizontal_line_valid.setter
    def is_horizontal_line_valid(self, value: bool) -> None:
        old_value = self.component.IsHorizontalLineValid
        if old_value != value:
            event_args = self._trigger_cancel_event("IsHorizontalLineValid", old_value, value)
            _ = self._trigger_done_event(event_args)

    @property
    def vertical_line(self) -> BorderLine2StructComp:
        """
        Gets/Set the vertical line style edge.

        Setting value can be done with a ``BorderLine2`` or ``BorderLine2Comp`` object.

        Returns:
            BorderLine2Comp: Returns Border Line.

        Hint:
            - ``BorderLine2`` can be imported from ``ooo.dyn.table.border_line2``
        """
        key = "vertical_line"
        prop = self._props.get(key, None)
        if prop is None:
            prop = BorderLine2StructComp(self.component.VerticalLine, key, self._events)
            self._props[key] = prop
        return cast(BorderLine2StructComp, prop)

    @vertical_line.setter
    def vertical_line(self, value: BorderLine2StructComp | BorderLine2) -> None:
        key = "vertical_line"
        old_value = self.component.VerticalLine
        if mInfo.Info.is_instance(value, BorderLine2StructComp):
            new_value = value.copy()
        else:
            comp = BorderLine2StructComp(cast(BorderLine2, value), key)
            new_value = comp.copy()

        event_args = self._trigger_cancel_event("VerticalLine", old_value, new_value)
        done_args = self._trigger_done_event(event_args)
        if done_args is None:
            return
        if key in self._props:
            del self._props[key]

    @property
    def is_vertical_line_valid(self) -> bool:
        """
        Gets/Sets whether the value of ``TableBorder2.VerticalLine`` is used.
        """
        return self.component.IsVerticalLineValid

    @is_vertical_line_valid.setter
    def is_vertical_line_valid(self, value: bool) -> None:
        old_value = self.component.IsVerticalLineValid
        if old_value != value:
            event_args = self._trigger_cancel_event("IsVerticalLineValid", old_value, value)
            _ = self._trigger_done_event(event_args)

    @property
    def distance(self) -> UnitMM100:
        """
        Gets/Sets the distance between the lines and other contents.

        When setting value it can be done with a ``int`` ( in ``1/100th mm`` units) or ``UnitT`` object.

        Returns:
            UnitMM100: Distance value.
        """
        return UnitMM100(self.component.Distance)  # type: ignore

    @distance.setter
    def distance(self, value: int | UnitT) -> None:
        val = UnitMM100.from_unit_val(value)
        new_val = val.value
        old_value = self.component.Distance
        if old_value != new_val:
            event_args = self._trigger_cancel_event("Distance", old_value, new_val)
            _ = self._trigger_done_event(event_args)

    @property
    def is_distance_valid(self) -> bool:
        """
        Gets/Sets whether the value of ``TableBorder2.Distance`` is used.
        """
        return self.component.IsVerticalLineValid

    @is_distance_valid.setter
    def is_distance_valid(self, value: bool) -> None:
        old_value = self.component.IsVerticalLineValid
        if old_value != value:
            event_args = self._trigger_cancel_event("IsVerticalLineValid", old_value, value)
            _ = self._trigger_done_event(event_args)

    # endregion Properties
