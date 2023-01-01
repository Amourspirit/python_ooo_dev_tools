# region imports
from __future__ import annotations
from typing import cast, overload

from ...exceptions import ex as mEx
from ...utils import props as mProps
from ..style_base import StyleBase
from . import side
from .side import Side as Side

import uno
from ooo.dyn.table.table_border import TableBorder
from ooo.dyn.table.table_border2 import TableBorder2


# endregion imports


class BorderTable(StyleBase):
    """Table Border positioning for use in styles."""

    _SIDE_ATTRS = ("TopLine", "BottomLine", "LeftLine", "RightLine", "HorizontalLine", "VerticalLine")

    # region init

    def __init__(
        self,
        left: Side | None = None,
        right: Side | None = None,
        top: Side | None = None,
        bottom: Side | None = None,
        border_side: Side | None = None,
        vertical: Side | None = None,
        horizontal: Side | None = None,
        distance: float | None = None,
    ) -> None:
        """
        Constructor

        Args:
            left (Side | None, optional): Determines the line style at the left edge.
            right (Side | None, optional): Determines the line style at the right edge.
            top (Side | None, optional): Determines the line style at the top edge.
            bottom (Side | None, optional): Determines the line style at the bottom edge.
            border_side (Side | None, optional): Determines the line style at the top, bottom, left, right edges. If this argument has a value then arguments ``top``, ``bottom``, ``left``, ``right`` are ignored
            horizontal (Side | None, optional): Determines the line style of horizontal lines for the inner part of a cell range.
            vertical (Side | None, optional): Determines the line style of vertical lines for the inner part of a cell range.
            distance (float | None, optional): Contains the distance between the lines and other contents (in mm units).
        """
        init_vals = {}
        if not border_side is None:
            init_vals["TopLine"] = border_side
            init_vals["IsTopLineValid"] = True
            init_vals["BottomLine"] = border_side
            init_vals["IsBottomLineValid"] = True
            init_vals["LeftLine"] = border_side
            init_vals["IsLeftLineValid"] = True
            init_vals["RightLine"] = border_side
            init_vals["IsRightLineValid"] = True

        else:
            if not top is None:
                init_vals["TopLine"] = top
                init_vals["IsTopLineValid"] = True
            if not bottom is None:
                init_vals["BottomLine"] = bottom
                init_vals["IsBottomLineValid"] = True
            if not left is None:
                init_vals["LeftLine"] = left
                init_vals["IsLeftLineValid"] = True
            if not right is None:
                init_vals["RightLine"] = right
                init_vals["IsRightLineValid"] = True

        if not horizontal is None:
            init_vals["HorizontalLine"] = horizontal
            init_vals["IsHorizontalLineValid"] = True
        if not vertical is None:
            init_vals["VerticalLine"] = vertical
            init_vals["IsVerticalLineValid"] = True
        if not distance is None:
            init_vals["Distance"] = round(distance * 100)
            init_vals["IsDistanceValid"] = True
        self._has_attribs = len(init_vals) > 0
        super().__init__(**init_vals)

    # endregion init

    # region methods

    # region apply_style()

    @overload
    def apply_style(self, obj: object) -> None:
        ...

    def apply_style(self, obj: object, **kwargs) -> None:
        """
        Applies Style to obj

        Args:
            obj (object): UNO object

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
        tb = cast(TableBorder2, mProps.Props.get(obj, "TableBorder2", None))
        if tb is None:
            raise mEx.PropertyNotFoundError("TableBorder2", "apply_style() obj has no property, TableBorder2")
        attrs = ("TopLine", "BottomLine", "LeftLine", "RightLine", "HorizontalLine", "VerticalLine")
        for attr in attrs:
            val = cast(Side, self._get(attr))
            if not val is None:
                setattr(tb, attr, val.get_border_line2())
                setattr(tb, f"Is{attr}Valid", True)
        distance = cast(int, self._get("Distance"))
        if not distance is None:
            tb.Distance = distance
            tb.IsDistanceValid = True
        mProps.Props.set(obj, TableBorder2=tb)

        h_line = cast(Side, self._get("HorizontalLine"))

        if h_line is None:
            h_ln = tb.HorizontalLine
            h_invalid = (
                h_ln.InnerLineWidth == 0
                and h_ln.LineDistance == 0
                and h_ln.LineWidth == 0
                and h_ln.OuterLineWidth == 0
            )
            if h_invalid:
                tb = cast(TableBorder2, mProps.Props.get(obj, "TableBorder2"))
                tb.HorizontalLine = Side.empty.get_border_line2()
                tb.IsHorizontalLineValid = True
                mProps.Props.set(obj, TableBorder2=tb)
        else:
            tb = cast(TableBorder2, mProps.Props.get(obj, "TableBorder2"))
            tb.HorizontalLine = h_line.get_border_line2()
            tb.IsHorizontalLineValid = True
            mProps.Props.set(obj, TableBorder2=tb)

    # endregion apply_style()

    @staticmethod
    def from_obj(obj: object) -> BorderTable:
        """
        Gets instance from object properties

        Args:
            obj (object): UNO object that has a ``TableBorder2`` property

        Raises:
            PropertyNotFoundError: If ``obj`` does not have ``TableBorder2`` property.

        Returns:
            BorderTable: Border Table.
        """
        tb = cast(TableBorder2, mProps.Props.get(obj, "TableBorder2", None))
        if tb is None:
            raise mEx.PropertyNotFoundError("TableBorder2", "from_obj() obj as no TableBorder2 property")
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
        bt = BorderTable(left=left, right=right, top=top, bottom=bottom, vertical=vertical, horizontal=horizontal)
        if tb.IsDistanceValid:
            bt._set("IsDistanceValid", True)
            bt._set("Distance", tb.Distance)
        bt._has_attribs = len(bt._dv) > 0
        return bt

    def get_table_border2(self) -> TableBorder2:
        """
        Gets Table Border for current instance.

        Returns:
            TableBorder2: ``com.sun.star.table.TableBorder2``
        """
        tb = TableBorder2()
        for key, val in self._dv.items():
            if key in BorderTable._SIDE_ATTRS:
                side = cast(Side, val)
                setattr(tb, key, side.get_border_line2())
            else:
                setattr(tb, key, val)
        return tb

    def get_table_border(self) -> TableBorder:
        """
        Gets Table Border for current instance.

        Returns:
            TableBorder: ``com.sun.star.table.TableBorder``
        """
        tb = TableBorder2()
        for key, val in self._dv.items():
            if key in BorderTable._SIDE_ATTRS:
                side = cast(Side, val)
                setattr(tb, key, side.get_border_line())
            else:
                setattr(tb, key, val)
        return tb

    # endregion methods

    # region Properties

    @property
    def distance(self) -> float | None:
        """Gets distance value"""
        pv = cast(int, self._get("Distance"))
        if not pv is None:
            if pv == 0:
                return 0.0
            return float(pv / 100)
        return None

    @property
    def left(self) -> Side | None:
        """Gets left value"""
        return self._get("LeftLine")

    @property
    def right(self) -> Side | None:
        """Gets right value"""
        return self._get("RightLine")

    @property
    def bottom(self) -> Side | None:
        """Gets bottom value"""
        return self._get("BottomLine")

    @property
    def horizontal(self) -> Side | None:
        """Gets horizontal value"""
        return self._get("HorizontalLine")

    @property
    def vertical(self) -> Side | None:
        """Gets vertical value"""
        return self._get("VerticalLine")

    @property
    def has_attribs(self) -> bool:
        """Gets If instantance has any attributes set."""
        return self._has_attribs

    # endregion Properties
