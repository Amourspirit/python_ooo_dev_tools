"""
Module for table border (``TableBorder2``) struct

.. versionadded:: 0.9.0
"""
# region imports
from __future__ import annotations
from typing import Tuple, Type, cast, overload, TypeVar

import uno
from ooo.dyn.table.table_border import TableBorder
from ooo.dyn.table.table_border2 import TableBorder2

from ....events.event_singleton import _Events
from ....exceptions import ex as mEx
from ....proto.unit_obj import UnitObj
from ....utils import props as mProps
from ....utils.unit_convert import UnitConvert, Length
from ...kind.format_kind import FormatKind
from ...style_base import StyleBase, EventArgs, CancelEventArgs, FormatNamedEvent
from ..common.props.struct_border_table_props import StructBorderTableProps
from ..common.props.prop_pair import PropPair
from .side import Side as Side


# endregion imports

_TTableBorderStruct = TypeVar(name="_TTableBorderStruct", bound="TableBorderStruct")


class TableBorderStruct(StyleBase):
    """
    Table Border struct positioning for use in styles.

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together Border Table properties.
    """

    # region init

    def __init__(
        self,
        *,
        left: Side | None = None,
        right: Side | None = None,
        top: Side | None = None,
        bottom: Side | None = None,
        border_side: Side | None = None,
        vertical: Side | None = None,
        horizontal: Side | None = None,
        distance: float | UnitObj | None = None,
    ) -> None:
        """
        Constructor

        Args:
            left (Side, optional): Determines the line style at the left edge.
            right (Side, optional): Determines the line style at the right edge.
            top (Side, optional): Determines the line style at the top edge.
            bottom (Side, optional): Determines the line style at the bottom edge.
            border_side (Side, optional): Determines the line style at the top, bottom, left, right edges. If this argument has a value then arguments ``top``, ``bottom``, ``left``, ``right`` are ignored
            horizontal (Side, optional): Determines the line style of horizontal lines for the inner part of a cell range.
            vertical (Side, optional): Determines the line style of vertical lines for the inner part of a cell range.
            distance (float, UnitObj, optional): Contains the distance between the lines and other contents (in mm units) or :ref:`proto_unit_obj`.
        """
        init_vals = {}
        if not border_side is None:
            init_vals[self._props.top.first] = border_side
            if self._props.top.second:
                init_vals[self._props.top.second] = True
            init_vals[self._props.bottom.first] = border_side
            if self._props.bottom.second:
                init_vals[self._props.bottom.second] = True
            init_vals[self._props.left.first] = border_side
            if self._props.left.second:
                init_vals[self._props.left.second] = True
            init_vals[self._props.right.first] = border_side
            if self._props.right.second:
                init_vals[self._props.right.second] = True
        else:
            if not top is None:
                init_vals[self._props.top.first] = top
                if self._props.top.second:
                    init_vals[self._props.top.second] = True
            if not bottom is None:
                init_vals[self._props.bottom.first] = bottom
                if self._props.bottom.second:
                    init_vals[self._props.bottom.second] = True
            if not left is None:
                init_vals[self._props.left.first] = left
                if self._props.left.second:
                    init_vals[self._props.left.second] = True
            if not right is None:
                init_vals[self._props.right.first] = right
                if self._props.right.second:
                    init_vals[self._props.right.second] = True

        if not horizontal is None:
            init_vals[self._props.horz.first] = horizontal
            if self._props.horz.second:
                init_vals[self._props.horz.second] = True
        if not vertical is None:
            init_vals[self._props.vert.first] = vertical
            if self._props.vert.second:
                init_vals[self._props.vert.second] = True
        if not distance is None:
            try:
                init_vals[self._props.dist.first] = distance.get_value_mm100()
            except AttributeError:
                init_vals[self._props.dist.first] = UnitConvert.convert(num=distance, frm=Length.MM, to=Length.MM100)
            if self._props.dist.second:
                init_vals[self._props.dist.second] = True
        super().__init__(**init_vals)

    # endregion init

    # region methods

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ()
        return self._supported_services_values

    def _get_property_name(self) -> str:
        try:
            return self._property_name
        except AttributeError:
            self._property_name = "TableBorder2"
        return self._property_name

    # region apply()

    @overload
    def apply(self, obj: object) -> None:
        ...

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies Style to obj

        Args:
            obj (object): UNO object

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_APPLYING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_APPLYED` :eventref:`src-docs-event`

        Raises:
            PropertyNotFoundError: If ``obj`` does not have ``TableBorder2`` property.

        Returns:
            None:
        """
        # there seems to be a bug in LibreOffice for TableBorder2.HorizontalLine
        # When top or bottom line is set for some reason TableBorder2.HorizontalLine also picks it up.
        #
        # current work around is to get if the current TableBorder2.HorizontalLine contains any data on read.
        # if it does not and the current style is not applying HorizontalLine then read back TableBorder2
        # and reset its HorizontalLine to default empty values.
        # the save TableBorder2 again.
        # Even if HorizontalLine is present in current style that is being applied it is ignored and set to top line or bottom line values.
        # The work around is to save TableBorder2 after setting other properties, read it again, set the HorizontalLine and save it again.
        if not self._is_valid_obj(obj):
            # will not apply on this class but may apply on child classes
            self._print_not_valid_obj("apply()")
            return

        prop_name = self._get_property_name()

        cargs = CancelEventArgs(source=f"{self.apply.__qualname__}")
        cargs.event_data = self
        self.on_applying(cargs)
        if cargs.cancel:
            return
        _Events().trigger(FormatNamedEvent.STYLE_APPLYING, cargs)
        if cargs.cancel:
            return

        tb = cast(TableBorder2, mProps.Props.get(obj, prop_name, None))
        if tb is None:
            raise mEx.PropertyNotFoundError(prop_name, "apply() obj has no property")
        attrs_props = (*(self._props[i] for i in range(6)),)

        for ap in attrs_props:
            val = cast(Side, self._get(ap.first))
            if not val is None:
                setattr(tb, ap.first, val.get_uno_struct())
                if ap.second:
                    setattr(tb, ap.second, True)
        distance = cast(int, self._get(self._props.dist.first))
        if not distance is None:
            tb.Distance = distance
            tb.IsDistanceValid = True
        mProps.Props.set(obj, TableBorder2=tb)

        h_line = cast(Side, self._get(self._props.horz.first))

        if h_line is None:
            h_ln = tb.HorizontalLine
            h_invalid = (
                h_ln.InnerLineWidth == 0
                and h_ln.LineDistance == 0
                and h_ln.LineWidth == 0
                and h_ln.OuterLineWidth == 0
            )
            if h_invalid:
                tb = cast(TableBorder2, mProps.Props.get(obj, prop_name))
                tb.HorizontalLine = Side.empty.get_uno_struct()
                tb.IsHorizontalLineValid = True
                mProps.Props.set(obj, TableBorder2=tb)
        else:
            tb = cast(TableBorder2, mProps.Props.get(obj, prop_name))
            tb.HorizontalLine = h_line.get_uno_struct()
            tb.IsHorizontalLineValid = True
            mProps.Props.set(obj, TableBorder2=tb)
        eargs = EventArgs.from_args(cargs)
        self.on_applied(eargs)
        _Events().trigger(FormatNamedEvent.STYLE_APPLIED, eargs)

    # endregion apply()

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TTableBorderStruct], obj: object) -> _TTableBorderStruct:
        ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TTableBorderStruct], obj: object, **kwargs) -> _TTableBorderStruct:
        ...

    @classmethod
    def from_obj(cls: Type[_TTableBorderStruct], obj: object, **kwargs) -> _TTableBorderStruct:
        """
        Gets instance from object properties

        Args:
            obj (object): UNO object

        Raises:
            PropertyNotFoundError: If ``obj`` does not have required property.

        Returns:
            BorderTable: Border Table.
        """
        # this nu is only used to get Property Name
        nu = cls(**kwargs)
        prop_name = nu._get_property_name()

        tb = cast(TableBorder2, mProps.Props.get(obj, prop_name, None))
        if tb is None:
            raise mEx.PropertyNotFoundError(prop_name, f"from_obj() obj as no {prop_name} property")
        line_props = ("Color", "InnerLineWidth", "LineDistance", "LineStyle", "LineWidth", "OuterLineWidth")

        left = Side() if tb.IsLeftLineValid else None
        top = Side() if tb.IsTopLineValid else None
        right = Side() if tb.IsRightLineValid else None
        bottom = Side() if tb.IsBottomLineValid else None
        vertical = Side() if tb.IsVerticalLineValid else None
        horizontal = Side() if tb.IsHorizontalLineValid else None

        for prop in line_props:
            if left:
                left._set(prop, getattr(tb.LeftLine, prop))
            if top:
                top._set(prop, getattr(tb.TopLine, prop))
            if right:
                right._set(prop, getattr(tb.RightLine, prop))
            if bottom:
                bottom._set(prop, getattr(tb.BottomLine, prop))
            if vertical:
                vertical._set(prop, getattr(tb.VerticalLine, prop))
            if horizontal:
                horizontal._set(prop, getattr(tb.HorizontalLine, prop))
        inst = cls(left=left, right=right, top=top, bottom=bottom, vertical=vertical, horizontal=horizontal, **kwargs)

        if tb.IsDistanceValid:
            p = inst._props.dist
            inst._set(p.first, tb.Distance)
            if p.second:
                inst._set(p.second, True)
        return inst

    # endregion from_obj()

    def get_uno_struct(self) -> TableBorder2:
        """
        Gets UNO ``TableBorder2`` from instance.

        Returns:
            TableBorder2: ``TableBorder2`` instance
        """
        tb = TableBorder2()

        # put attribs in a tuple
        attrs = (*(self._props[i].first for i in range(6)),)

        for key, val in self._dv.items():
            if key in attrs:
                side = cast(Side, val)
                setattr(tb, key, side.get_uno_struct())
            else:
                setattr(tb, key, val)
        return tb

    def get_uno_struct_table_border(self) -> TableBorder:
        """
        Gets UNO ``TableBorder`` from instance.

        Returns:
            TableBorder: ``TableBorder`` instance
        """
        tb = TableBorder2()
        # put attribs in a tuple
        attrs = (*(self._props[i].first for i in range(6)),)

        for key, val in self._dv.items():
            if key in attrs:
                side = cast(Side, val)
                setattr(tb, key, side.get_uno_struct_border_line())
            else:
                setattr(tb, key, val)
        return tb

    # endregion methods

    # region Style methods
    def fmt_border_side(self: _TTableBorderStruct, value: Side | None) -> _TTableBorderStruct:
        """
        Gets copy of instance with left, right, top, bottom sides set or removed

        Args:
            value (Side | None): Side value

        Returns:
            BorderTable: Border Table
        """
        cp = self.copy()
        cp.prop_top = value
        cp.prop_bottom = value
        cp.prop_left = value
        cp.prop_right = value
        return cp

    def fmt_top(self: _TTableBorderStruct, value: Side | None) -> _TTableBorderStruct:
        """
        Gets a copy of instance with top side set or removed

        Args:
            value (Side | None): Side value

        Returns:
            BorderTable: Border Table
        """
        cp = self.copy()
        cp.prop_top = value
        return cp

    def fmt_bottom(self: _TTableBorderStruct, value: Side | None) -> _TTableBorderStruct:
        """
        Gets a copy of instance with bottom side set or removed

        Args:
            value (Side | None): Side value

        Returns:
            BorderTable: Border Table
        """
        cp = self.copy()
        cp.prop_bottom = value
        return cp

    def fmt_left(self: _TTableBorderStruct, value: Side | None) -> _TTableBorderStruct:
        """
        Gets a copy of instance with left side set or removed

        Args:
            value (Side | None): Side value

        Returns:
            BorderTable: Border Table
        """
        cp = self.copy()
        cp.prop_left = value
        return cp

    def fmt_right(self: _TTableBorderStruct, value: Side | None) -> _TTableBorderStruct:
        """
        Gets a copy of instance with right side set or removed

        Args:
            value (Side | None): Side value

        Returns:
            BorderTable: Border Table
        """
        cp = self.copy()
        cp.prop_right = value
        return cp

    def fmt_horizontal(self: _TTableBorderStruct, value: Side | None) -> _TTableBorderStruct:
        """
        Gets a copy of instance with horizontal side set or removed

        Args:
            value (Side | None): Side value

        Returns:
            BorderTable: Border Table
        """
        cp = self.copy()
        cp.prop_horizontal = value
        return cp

    def fmt_vertical(self: _TTableBorderStruct, value: Side | None) -> _TTableBorderStruct:
        """
        Gets a copy of instance with top vertical set or removed

        Args:
            value (Side | None): Side value

        Returns:
            BorderTable: Border Table
        """
        cp = self.copy()
        cp.prop_vertical = value
        return cp

    def fmt_distance(self: _TTableBorderStruct, value: float | None) -> _TTableBorderStruct:
        """
        Gets a copy of instance with distance set or removed

        Args:
            value (float | None): Distance value

        Returns:
            BorderTable: Border Table
        """
        cp = self.copy()
        cp.prop_distance = value
        return cp

    # endregion Style methods

    # region Properties

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.STRUCT
        return self._format_kind_prop

    @property
    def prop_distance(self) -> float | None:
        """Gets/Sets distance value (``in mm`` units)"""
        pv = cast(int, self._get(self._props.dist.first))
        if not pv is None:
            if pv == 0:
                return 0.0
            return UnitConvert.convert(num=pv, frm=Length.MM100, to=Length.MM)
        return None

    @prop_distance.setter
    def prop_distance(self, value: float | UnitObj | None) -> None:
        p = self._props.dist
        if value is None:
            self._remove(p.first)
            if p.second:
                self._remove(p.second)
            return
        try:
            self._set(p.first, value.get_value_mm100())
        except AttributeError:
            self._set(p.first, UnitConvert.convert(num=value, frm=Length.MM, to=Length.MM100))
        if p.second:
            self._set(p.second, True)

    @property
    def prop_left(self) -> Side | None:
        """Gets/Sets left value"""
        return self._get(self._props.left.first)

    @prop_left.setter
    def prop_left(self, value: Side | None) -> None:
        p = self._props.left
        if value is None:
            self._remove(p.first)
            if p.second:
                self._remove(p.second)
            return
        self._set(p.first, value)
        if p.second:
            self._set(p.second, True)

    @property
    def prop_right(self) -> Side | None:
        """Gets/Sets right value"""
        return self._get(self._props.right.first)

    @prop_right.setter
    def prop_right(self, value: Side | None) -> None:
        p = self._props.right
        if value is None:
            self._remove(p.first)
            if p.second:
                self._remove(p.second)
            return
        self._set(p.first, value)
        if p.second:
            self._set(p.second, True)

    @property
    def prop_top(self) -> Side | None:
        """Gets/Sets bottom value"""
        return self._get(self._props.top.first)

    @prop_top.setter
    def prop_top(self, value: Side | None) -> None:
        p = self._props.top
        if value is None:
            self._remove(p.first)
            if p.second:
                self._remove(p.second)
            return
        self._set(p.first, value)
        if p.second:
            self._set(p.second, True)

    @property
    def prop_bottom(self) -> Side | None:
        """Gets/Sets bottom value"""
        return self._get(self._props.bottom.first)

    @prop_bottom.setter
    def prop_bottom(self, value: Side | None) -> None:
        p = self._props.bottom
        if value is None:
            self._remove(p.first)
            if p.second:
                self._remove(p.second)
            return
        self._set(p.first, value)
        if p.second:
            self._set(p.second, True)

    @property
    def prop_horizontal(self) -> Side | None:
        """Gets/Sets horizontal value"""
        return self._get(self._props.horz.first)

    @prop_horizontal.setter
    def prop_horizontal(self, value: Side | None) -> None:
        p = self._props.horz
        if value is None:
            self._remove(p.first)
            if p.second:
                self._remove(p.second)
            return
        self._set(p.first, value)
        if p.second:
            self._set(p.second, True)

    @property
    def prop_vertical(self) -> Side | None:
        """Gets/Sets vertical value"""
        return self._get(self._props.vert.first)

    @prop_vertical.setter
    def prop_vertical(self, value: Side | None) -> None:
        p = self._props.vert
        if value is None:
            self._remove(p.first)
            if p.second:
                self._remove(p.second)
            return
        self._set(p.first, value)
        if p.second:
            self._set(p.second, True)

    @property
    def _props(self) -> StructBorderTableProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = StructBorderTableProps(
                left=PropPair("LeftLine", "IsLeftLineValid"),
                top=PropPair("TopLine", "IsTopLineValid"),
                right=PropPair("RightLine", "IsRightLineValid"),
                bottom=PropPair("BottomLine", "IsBottomLineValid"),
                horz=PropPair("HorizontalLine", "IsHorizontalLineValid"),
                vert=PropPair("VerticalLine", "IsVerticalLineValid"),
                dist=PropPair("Distance", "IsDistanceValid"),
            )
        return self._props_internal_attributes

    # endregion Properties
