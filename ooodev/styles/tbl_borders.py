# region imports
from __future__ import annotations
from typing import cast, overload
from enum import IntFlag

from . import paragraphs as mPara
from ..exceptions import ex as mEx
from ..meta.static_prop import static_prop
from ..utils import info as mInfo
from ..utils import lo as mLo
from ..utils import props as mProps
from ..utils.color import Color
from ..utils.color import CommonColor
from .style_base import StyleBase
from .style_const import POINT_RATIO

import uno
from ooo.dyn.table.border_line import BorderLine as BorderLine
from ooo.dyn.table.border_line_style import BorderLineStyleEnum as BorderLineStyleEnum
from ooo.dyn.table.border_line2 import BorderLine2 as BorderLine2
from ooo.dyn.table.shadow_format import ShadowFormat as ShadowFormat
from ooo.dyn.table.shadow_location import ShadowLocation as ShadowLocation
from ooo.dyn.table.table_border import TableBorder
from ooo.dyn.table.table_border2 import TableBorder2


# endregion imports

# region Enums


class BorderKind(IntFlag):
    """Border Value (bitwise composition is possible)"""

    TOP = 0x01
    """Apply Top Border"""
    BOTTOM = 0x02
    """Apply Bottom Border"""
    LEFT = 0x04
    """Apply Left Border"""
    RIGHT = 0x08
    """Apply Right Border"""
    ALL = TOP | BOTTOM | LEFT | RIGHT
    """Apply to all Borders"""


class SideFlags(IntFlag):
    """Side Flags Enum"""

    LEFT = 0x01
    """Apply to Left Side"""
    TOP = 0x02
    """Apply to Top Side"""
    RIGHT = 0x03
    """Apply to Right Side"""
    BOTTOM = 0x04
    """Apply to Bottom Side"""
    LEFT_RIGHT = LEFT | RIGHT
    """Apply to Left and Right Sides"""
    TOP_BOTTOM = TOP | BOTTOM
    """Apply to Top and Bottom Sides"""
    BORDER = LEFT | TOP | RIGHT | BOTTOM
    """Apply to Left, Right, Top, and bottom Sides"""
    BOTTOM_LEFT_TOP_RIGHT = 0x05
    """Apply to Diagonal starting Bottom-Left and draw to Top-Right"""
    TOP_LEFT_BOTTOM_RIGHT = 0x06
    """Apply to Diagonal starting Top-Left and draw to Bottom-Right"""


# endregion Enums


class Side(StyleBase):
    """Represents one side of a border"""

    _EMPTY = None

    # region init

    def __init__(
        self,
        style: BorderLineStyleEnum = BorderLineStyleEnum.SOLID,
        color: Color = CommonColor.BLACK,
        width: float = 0.75,
        width_inner: float = 0.0,
        distance: float = 0.0,
    ) -> None:
        """
        Constructs Side

        Args:
            style (BorderLineStyleEnum, optional): Line Style of the border.
            color (Color, optional): Color of the border.
            width (float, optional): Contains the width in of a single line or the width of outer part of a double line (in mm units). If this value is zero, no line is drawn. Default ``0.26``
            width_inner (float, optional): contains the width of the inner part of a double line (in mm units). If this value is zero, only a single line is drawn. Default ``0.0``
            distance (float, optional): contains the distance between the inner and outer parts of a double line (in mm units). Defalut ``0.0``

        Raises:
            ValueError: if ``color``, ``width`` or ``width_inner`` is less than ``0``.

        Returns:
            None:
        """
        if color < 0:
            raise ValueError("color must be a positive value")
        if width < 0.0:
            raise ValueError("width must be a postivie value")
        if width_inner < 0.0:
            raise ValueError("width_inner must be a postivie value")
        if distance < 0.0:
            raise ValueError("distance must be a postivie value")

        self._style = style
        self._color = color
        self._width = width
        self._width_inner = width_inner
        self._distance = distance

        init_vals = {
            "Color": self._color,
            "InnerLineWidth": round(self._width_inner * POINT_RATIO),
            "LineDistance": round(self._distance * 100),
            "LineStyle": self._style.value,
            "LineWidth": round(self._width * POINT_RATIO),
            "OuterLineWidth": round(self._width * POINT_RATIO),
        }

        super().__init__(**init_vals)

    # endregion init

    # region methods

    @overload
    def apply_style(self, obj: object, *, flags: SideFlags) -> None:
        ...

    def apply_style(self, obj: object, **kwargs) -> None:
        """
        Applies style to object

        Args:
            obj (object): Object to apply style to.

        Other Parameters:
            flags: (SideFlags): Determins where to apply side.

        Returns:
            None:
        """
        expected_key_names = ("flags",)
        for kw in expected_key_names:
            if not kw in kwargs:
                raise Exception(f'apply_style() requires argument "{kw}"')
        flags = cast(SideFlags, kwargs["flags"])
        val = self.get_border_line2()
        if SideFlags.LEFT in flags:
            mProps.Props.set(obj, LeftBorder2=val)
        if SideFlags.TOP in flags:
            mProps.Props.set(obj, TopBorder2=val)
        if SideFlags.RIGHT in flags:
            mProps.Props.set(obj, RightBorder2=val)
        if SideFlags.BOTTOM in flags:
            mProps.Props.set(obj, BottomBorder2=val)
        if SideFlags.BOTTOM_LEFT_TOP_RIGHT in flags:
            mProps.Props.set(obj, DiagonalBLTR2=val)
        if SideFlags.TOP_LEFT_BOTTOM_RIGHT in flags:
            mProps.Props.set(obj, DiagonalTLBR2=val)

    def get_border_line(self) -> BorderLine:
        b2 = self.get_border_line2()
        return BorderLine(
            Color=b2.Color,
            InnerLineWidth=b2.InnerLineWidth,
            OuterLineWidth=b2.OuterLineWidth,
            LineDistance=b2.LineDistance,
        )

    def get_border_line2(self) -> BorderLine2:
        """gets Border Line of instance"""
        line = BorderLine2()  # create the border line
        for key, val in self._dv.items():
            setattr(line, key, val)
        return line

    @static_prop
    def empty(cls) -> Side:
        """Gets an empyty side. When applied formatting is removed"""
        if cls._EMPTY is None:
            cls._EMPTY = Side(style=BorderLineStyleEnum.NONE, color=0, width=0.0, width_inner=0.0, distance=0)
        return cls._EMPTY

    # endregion methods

    # region properties

    @property
    def style(self) -> BorderLineStyleEnum:
        """Gets Border Line style"""
        return self._style

    @property
    def color(self) -> Color:
        """Gets Border Line Color"""
        return self._color

    @property
    def width(self) -> float:
        """
        Gets Border Line Width.

        Contains the width of a single line or the width of outer part of a double line (in mm units).
        If this value is zero, no line is drawn.
        """
        return self._width

    @property
    def width_inner(self) -> float:
        """
        Gets Border Line Inner Width.

        Contains the width of the inner part of a double line (in mm units).
        If this value is zero, no line is drawn.
        """
        return self._width_inner

    # endregion properties


class Shadow(StyleBase):
    # region init
    _EMPTY = None

    def __init__(
        self,
        location: ShadowLocation = ShadowLocation.BOTTOM_RIGHT,
        color: Color = CommonColor.GRAY,
        transparent: bool = False,
        width: float = 1.76,
    ) -> None:
        """
        Constructor

        Args:
            location (ShadowLocation, optional): contains the location of the shadow. Default to ``ShadowLocation.BOTTOM_RIGHT``.
            color (Color, optional):contains the color value of the shadow. Defaults to ``CommonColor.GRAY``.
            transparent (bool, optional): Shadow transparency. Defaults to False.
            width (float, optional): contains the size of the shadow (in mm units). Defaults to ``1.76``.

        Raises:
            ValueError: If ``color`` or ``width`` are less than zero.
        """
        if color < 0:
            raise ValueError("color must be a positive number")
        if width < 0:
            raise ValueError("Width must be a postivie number")
        init_vals = {
            "Location": location,
            "Color": color,
            "IsTransparent": transparent,
            "ShadowWidth": round(width * 100),
        }
        super().__init__(**init_vals)

    # endregion init

    # region methods

    def get_shadow_format(self) -> ShadowFormat:
        """
        Gets Shadow format for instance.

        Returns:
            ShadowFormat: Shadow Format
        """
        return ShadowFormat(
            Location=self._get("Location"),
            ShadowWidth=self._get("ShadowWidth"),
            IsTransparent=self._get("IsTransparent"),
            Color=self._get("Color"),
        )

    # region apply_style()

    @overload
    def apply_style(self, obj: object) -> None:
        ...

    def apply_style(self, obj: object, **kwargs) -> None:
        """
        Applies style to object

        Args:
            obj (object): Object that contains a ``ShadowFormat`` property.

        Returns:
            None:
        """
        shadow = self.get_shadow_format()
        mProps.Props.set(obj, ShadowFormat=shadow)

    # endregion apply_style()

    # endregion methods

    # region Properties

    @property
    def location(self) -> ShadowLocation:
        """Gets the location of the shadow."""
        return self._get("Location")

    @property
    def color(self) -> Color:
        """Gets the color value of the shadow."""
        return self._get("Color")

    @property
    def transparent(self) -> bool:
        """Gets transparent value"""
        return self._get("IsTransparent")

    @property
    def width(self) -> float:
        """Gets the size of the shadow (in mm units)"""
        pv = cast(int, self._get("ShadowWidth"))
        if pv == 0:
            return 0.0
        return float(pv / 100)

    @static_prop
    def empty(cls) -> Shadow:
        """Gets empty Shadow. Static Property. when style is applied it remove any shadow."""
        if cls._EMPTY is None:
            cls._EMPTY = Shadow(location=ShadowLocation.NONE, transparent=False, color=8421504)
            # just to be exact due to float conversions.
            cls._EMPTY._set("ShadowWidth", 176)
        return cls._EMPTY

    # endregion Properties


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


class Border(StyleBase):
    """Border positioning for use in styles."""

    _DEFAULT = None
    _EMPTY = None

    # region init

    def __init__(
        self,
        *,
        right: Side | None = None,
        left: Side | None = None,
        top: Side | None = None,
        bottom: Side | None = None,
        border_side: Side | None = None,
        vertical: Side | None = None,
        horizontal: Side | None = None,
        distance: float | None = None,
        diagonal_down: Side | None = None,
        diagonal_up: Side | None = None,
        shadow: Shadow | None = None,
        padding: mPara.Padding | None = None,
    ) -> None:
        """
        _summary_

        Args:
            left (Side | None, optional): Determines the line style at the left edge.
            right (Side | None, optional): Determines the line style at the right edge.
            top (Side | None, optional): Determines the line style at the top edge.
            bottom (Side | None, optional): Determines the line style at the bottom edge.
            border_side (Side | None, optional): Determines the line style at the top, bottom, left, right edges. If this argument has a value then arguments ``top``, ``bottom``, ``left``, ``right`` are ignored
            horizontal (Side | None, optional): Determines the line style of horizontal lines for the inner part of a cell range.
            vertical (Side | None, optional): Determines the line style of vertical lines for the inner part of a cell range.
            distance (float | None, optional): Contains the distance between the lines and other contents (in mm units).
            diagonal_down (Side | None, optional): Determines the line style from top-left to bottom-right diagonal.
            diagonal_up (Side | None, optional): Determines the line style from bottom-left to top-right diagonal.
            shadow (Shadow | None, optional): Cell Shadow
            padding (BorderPadding | None, optional): Cell padding
        """
        init_vals = {}

        if not shadow is None:
            init_vals["ShadowFormat"] = shadow.get_shadow_format()
        if not diagonal_down is None:
            init_vals["DiagonalTLBR2"] = diagonal_down.get_border_line2()
        if not diagonal_up is None:
            init_vals["DiagonalBLTR2"] = diagonal_up.get_border_line2()

        border_table = BorderTable(
            left=left,
            right=right,
            top=top,
            bottom=bottom,
            border_side=border_side,
            horizontal=horizontal,
            vertical=vertical,
            distance=distance,
        )

        if border_table.has_attribs:
            self._border_table = border_table
        else:
            self._border_table = None
        self._padding = padding

        super().__init__(**init_vals)

    # endregion init

    # region methods

    def apply_style(self, obj: object, **kwargs) -> None:
        """
        Applies padding to ``obj``

        Args:
            obj (object): Object that supports ``com.sun.star.style.ParagraphProperties`` service.
            kwargs (Any, optional): Expandable list of key value pairs that may be used in child classes.

        Returns:
            None:
        """
        if not self._padding is None:
            self._padding.apply_style(obj)
        if not self._border_table is None:
            self._border_table.apply_style(obj)
        if mInfo.Info.support_service(obj, "com.sun.star.table.CellProperties"):
            try:
                super().apply_style(obj)
            except mEx.MultiError as e:
                mLo.Lo.print(f"CellBorder.apply_style(): Unable to set Property")
                for err in e.errors:
                    mLo.Lo.print(f"  {err}")

    # endregion methods

    # region Properties
    @static_prop
    def default(cls) -> Border:
        """Gets Default Border. Static Property"""
        if cls._DEFAULT is None:
            cls._DEFAULT = Border(border_side=Side(), padding=mPara.Padding.default)
        return cls._DEFAULT

    @static_prop
    def empty(cls) -> Border:
        """Gets Empty Border. Static Property. When style is applied formatting is removed."""
        if cls._EMPTY is None:
            cls._EMPTY = Border(
                border_side=Side.empty,
                vertical=Side.empty,
                horizontal=Side.empty,
                diagonal_down=Side.empty,
                diagonal_up=Side.empty,
                distance=0.0,
                shadow=Shadow.empty,
                padding=mPara.Padding.default,
            )
        return cls._EMPTY

    # endregion Properties
