from __future__ import annotations
from typing import Any, cast
from enum import Enum, IntFlag

from ..exceptions import ex as mEx
from ..utils import info as mInfo
from ..utils import lo as mLo
from ..utils import props as mProps
from ..utils.color import Color
from ..utils.color import CommonColor
from .style_base import StyleBase
from . import paragraphs as mPara
from ..events.args.key_val_cancel_args import KeyValCancelArgs

import uno
from ooo.dyn.table.border_line_style import BorderLineStyleEnum as LineStyleKind
from ooo.dyn.table.shadow_location import ShadowLocation as ShadowLocation
from ooo.dyn.table.shadow_format import ShadowFormat as ShadowFormat
from ooo.dyn.table.border_line import BorderLine as BorderLine
from ooo.dyn.table.border_line2 import BorderLine2 as BorderLine2
from ooo.dyn.table.table_border2 import TableBorder2
from ooo.dyn.table.table_border import TableBorder


class BorderKind(IntFlag):
    """Border Value (bitwise composition is possible)"""

    TOP_BORDER = 0x01
    """Apply Top Border"""
    BOTTOM_BORDER = 0x02
    """Apply Bottom Border"""
    LEFT_BORDER = 0x04
    """Apply Left Border"""
    RIGHT_BORDER = 0x08
    """Apply Right Border"""
    ALL = TOP_BORDER | BOTTOM_BORDER | LEFT_BORDER | RIGHT_BORDER
    """Apply to all Borders"""


class Side:
    """Represents one side of a border"""

    def __init__(
        self,
        style: LineStyleKind = LineStyleKind.SOLID,
        color: Color = CommonColor.BLACK,
        width: float = 0.26,
        width_inner: float = 0.0,
        distance: float = 0.0,
    ) -> None:
        """
        Constructs Side

        Args:
            style (LineStyleKind, optional): Line Style of the border.
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

        init_vals = {
            "Color": color,
            "InnerLineWidth": round(width_inner * 100),
            "LineDistance": round(distance * 100),
            "LineStyle": style.value,
            "LineWidth": round(width * 100),
            "OuterLineWidth": round(width * 100),
        }
        self._dv = init_vals

    def _get(self, key: str) -> Any:
        return self._dv.get(key, None)

    @property
    def style(self) -> LineStyleKind:
        """Gets Border Line style"""
        pv = cast(int, self._get("LineStyle"))
        return LineStyleKind(pv)

    @property
    def color(self) -> Color:
        """Gets Border Line Color"""
        return self._get("Color")

    @property
    def width(self) -> float:
        """
        Gets Border Line Width.

        Contains the width of a single line or the width of outer part of a double line (in mm units).
        If this value is zero, no line is drawn.
        """
        pv = cast(int, self._get("OuterLineWidth"))
        if pv == 0:
            return 0.0
        return float(pv / 100)

    @property
    def width_inner(self) -> float:
        """
        Gets Border Line Inner Width.

        Contains the width of the inner part of a double line (in mm units).
        If this value is zero, no line is drawn.
        """
        pv = cast(int, self._get("InnerLineWidth"))
        if pv == 0:
            return 0.0
        return float(pv / 100)

    # def apply_style(self, obj: object, border: BorderKind = BorderKind.ALL) -> None:
    #     tb = cast(TableBorder, mProps.Props.get(obj, "TableBorder", None))
    #     tb2 = cast(TableBorder2, mProps.Props.get(obj, "TableBorder2", None))
    #     if tb is None and tb2 is None:
    #         mLo.Lo.print("Side.apply_style(): No border to apply styles")
    #         return
    #     bl = self.get_border_line()
    #     bl2 = self.get_border_line2()
    #     if BorderKind.TOP_BORDER in border:
    #         if not tb is None:
    #             tb.TopLine = bl
    #             tb.IsTopLineValid = True
    #         if not tb2 is None:
    #             tb2.TopLine = bl2
    #             tb2.IsTopLineValid = True
    #     if BorderKind.BOTTOM_BORDER in border:
    #         if not tb is None:
    #             tb.BottomLine = bl
    #             tb.IsBottomLineValid = True
    #         if not tb2 is None:
    #             tb2.BottomLine = bl2
    #             tb2.IsBottomLineValid = True
    #     if BorderKind.LEFT_BORDER in border:
    #         if not tb is None:
    #             tb.LeftLine = bl
    #             tb.IsLeftLineValid = True
    #         if not tb2 is None:
    #             tb2.LeftLine = bl2
    #             tb2.IsLeftLineValid = True
    #     if BorderKind.RIGHT_BORDER in border:
    #         if not tb is None:
    #             tb.RightLine = bl
    #             tb.IsRightLineValid = True
    #         if not tb2 is None:
    #             tb2.RightLine = bl2
    #             tb2.IsRightLineValid = True

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


class Shadow:
    def __init__(
        self, location: ShadowLocation = ShadowLocation.BOTTOM_RIGHT, color: Color = CommonColor.GRAY, transparent: bool = False, width: float = 1.76
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
        self._color = color
        self._transparent = transparent
        self._width = round(width * 100)
        self._location = location

    def get_shadow_format(self) -> ShadowFormat:
        return ShadowFormat(
            Location=self._location, ShadowWidth=self._width, IsTransparent=self._transparent, Color=self._color
        )

    @property
    def location(self) -> ShadowLocation:
        """Gets the location of the shadow."""
        return self._location

    @property
    def color(self) -> Color:
        """Gets the color value of the shadow."""
        return self._color

    @property
    def transparent(self) -> bool:
        """Gets transparent value"""
        return self._transparent

    @property
    def width(self) -> float:
        """Gets the size of the shadow (in mm units)"""
        if self._width == 0.0:
            return 0.0
        return float(self._width / 100)


class BorderPadding(mPara.Padding):
    def __init__(
        self,
        left: float = 0.35,
        right: float = 0.35,
        top: float = 0.35,
        bottom: float = 0.35,
        padding_all: float | None = None,
    ) -> None:
        """
        Constructor

        Args:
            left (float, optional): Paragraph left padding (in mm units). Defaults to ``0.35``.
            right (float, optional): Paragraph right padding (in mm units). Defaults to ``0.35``.
            top (float, optional): Paragraph top padding (in mm units). Defaults to ``0.35``.
            bottom (float, optional): Paragraph bottom padding (in mm units). Defaults to ``0.35``.
            padding_all (float, optional): Paragraph left, right, top, bottom padding (in mm units). If argument is present then ``left``, ``right``, ``top``, and ``bottom`` arguments are ignored.
        """
        super().__init__(left, right, top, bottom, padding_all)

    @property
    def left(self) -> float:
        """Gets paragraph left padding (in mm units)."""
        return super().left

    @property
    def right(self) -> float:
        """Gets paragraph right padding (in mm units)."""
        return super().right

    @property
    def top(self) -> float:
        """Gets paragraph top padding (in mm units)."""
        return super().top

    @property
    def bottom(self) -> float:
        """Gets paragraph bottom padding (in mm units)."""
        return super().bottom


class BorderTable:
    """Table Border positioning for use in styles."""

    _SIDE_ATTRS = ("TopLine", "BottomLine", "LeftLine", "RightLine", "HorizontalLine", "VerticalLine")

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
        self._dv = init_vals

    def _get(self, key: str) -> Any:
        return self._dv.get(key, None)

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


class CellBorder(StyleBase):
    """Border positioning for use in styles."""

    def __init__(
        self,
        *,
        top: Side | None = None,
        bottom: Side | None = None,
        left: Side | None = None,
        right: Side | None = None,
        border_side: Side | None = None,
        vertical: Side | None = None,
        horizontal: Side | None = None,
        distance: float | None = None,
        diagonal_down: Side | None = None,
        diagonal_up: Side | None = None,
        shadow: Shadow | None = None,
        padding: BorderPadding | None = None,
    ) -> None:
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
            init_vals["TableBorder2"] = border_table.get_table_border2()
        self._padding = padding

        super().__init__(**init_vals)

    def apply_style(self, obj: object) -> None:
        if not self._padding is None:
            self._padding.apply_style(obj)
        if mInfo.Info.support_service(obj, "com.sun.star.table.CellProperties"):
            try:
                super().apply_style(obj)
            except mEx.MultiError as e:
                mLo.Lo.print(f"CellBorder.apply_style(): Unable to set Property")
                for err in e.errors:
                    mLo.Lo.print(f"  {err}")
