"""
Module for table side (``BorderLine2``) struct.

.. versionadded:: 0.9.0
"""
# region imports
from __future__ import annotations
from typing import Dict, Tuple, cast, overload
from enum import IntFlag, Enum

from ....events.event_singleton import _Events
from ....meta.static_prop import static_prop
from ....utils import props as mProps
from ....utils.color import Color
from ....utils.color import CommonColor
from ...kind.format_kind import FormatKind
from ...style_base import StyleBase, EventArgs, CancelEventArgs, FormatNamedEvent
from ...style_const import POINT_RATIO
from ..common import border_width_impl as mBwi
from ....utils.unit_convert import UnitConvert, Length

import uno
from ooo.dyn.table.border_line import BorderLine as BorderLine
from ooo.dyn.table.border_line_style import BorderLineStyleEnum as BorderLineStyleEnum
from ooo.dyn.table.border_line2 import BorderLine2 as BorderLine2


# endregion imports

# region Enums


class LineSize(Enum):
    """Line Size Options"""

    HAIRLINE = (1, 0.05)
    """``0.05pt``"""
    VERY_THIN = (2, 0.5)
    """``0.5pt``"""
    THIN = (3, 0.75)
    """``0.75pt``"""
    MEDIUM = (4, 1.5)
    """``1.5pt``"""
    THICK = (4, 2.25)
    """``2.5pt``"""
    EXTRA_THICK = (5, 4.5)
    """``5.5pt``"""

    def __float__(self) -> float:
        return self.value[1]


class SideFlags(IntFlag):
    """
    Side Flags Enum

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

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
    """
    Side struct.

    Represents one side of a border.

    .. versionadded:: 0.9.0
    """

    _EMPTY = None

    # region init

    def __init__(
        self,
        line: BorderLineStyleEnum = BorderLineStyleEnum.SOLID,
        color: Color = CommonColor.BLACK,
        width: LineSize | float = LineSize.THIN,
    ) -> None:
        """
        Constructs Side

        Args:
            line (BorderLineStyleEnum, optional): Line Style of the border.
            color (Color, optional): Color of the border.
            width (LineSize, float, optional): Contains the width in of a single line or the width of outer part of a double line (in pt units). If this value is zero, no line is drawn. Default ``0.75``

        Raises:
            ValueError: if ``color``, ``width`` or ``width_inner`` is less than ``0``.

        Returns:
            None:
        """
        width = float(width)
        if color < 0:
            raise ValueError("color must be a positive value")
        if width < 0.0:
            raise ValueError("width must be a postivie value")
        if width > 9.0000001:
            raise ValueError("Maximum width allowed is 9pt")

        self._pts = width

        lw = round(UnitConvert.convert(num=width, frm=Length.PT, to=Length.MM100))

        init_vals = {
            "Color": color,
            "InnerLineWidth": 0,
            "LineDistance": 0,
            "LineStyle": line.value,
            "LineWidth": lw,
            "OuterLineWidth": 0,
        }

        super().__init__(**init_vals)
        self._set_line_values(pts=width, line=line)

    def _set_line_values(self, pts: int, line: BorderLineStyleEnum) -> None:
        if line == BorderLineStyleEnum.BORDER_LINE_STYLE_MAX:
            raise ValueError("BORDER_LINE_STYLE_MAX is not supported")

        val_keys = ("OuterLineWidth", "InnerLineWidth", "LineDistance")

        if line == BorderLineStyleEnum.NONE:
            for attr in val_keys:
                self._set(attr, 0)
            return

        twips = UnitConvert.to_twips(pts, Length.PT)
        single_lns = (
            BorderLineStyleEnum.SOLID,
            BorderLineStyleEnum.DOTTED,
            BorderLineStyleEnum.DASHED,
            BorderLineStyleEnum.FINE_DASHED,
            BorderLineStyleEnum.DASH_DOT,
            BorderLineStyleEnum.DASH_DOT_DOT,
        )
        en_em = (BorderLineStyleEnum.ENGRAVED, BorderLineStyleEnum.EMBOSSED)

        vals = None

        if line in single_lns:
            flags = mBwi.BorderWidthImplFlags.CHANGE_LINE1
            bw = mBwi.BorderWidthImpl(nFlags=flags, nRate1=1.0, nRate2=0.0, nRateGap=0.0)
            vals = (bw.get_line1(twips), bw.get_line2(twips), bw.get_gap(twips))

        elif line == BorderLineStyleEnum.DOUBLE:
            flags = (
                mBwi.BorderWidthImplFlags.CHANGE_DIST
                | mBwi.BorderWidthImplFlags.CHANGE_LINE1
                | mBwi.BorderWidthImplFlags.CHANGE_LINE2
            )
            bw = mBwi.BorderWidthImpl(nFlags=flags, nRate1=1 / 3, nRate2=1 / 3, nRateGap=1 / 3)
            vals = (bw.get_line1(twips), bw.get_line2(twips), bw.get_gap(twips))

        elif line == BorderLineStyleEnum.DOUBLE_THIN:
            flags = mBwi.BorderWidthImplFlags.CHANGE_DIST
            bw = mBwi.BorderWidthImpl(nFlags=flags, nRate1=10.0, nRate2=10.0, nRateGap=1.0)
            vals = (bw.get_line1(twips), bw.get_line2(twips), bw.get_gap(twips))

        elif line == BorderLineStyleEnum.THINTHICK_SMALLGAP:
            flags = mBwi.BorderWidthImplFlags.CHANGE_LINE1
            bw = mBwi.BorderWidthImpl(
                nFlags=flags, nRate1=1.0, nRate2=mBwi.THINTHICK_SMALLGAP_LINE2, nRateGap=mBwi.THINTHICK_SMALLGAP_GAP
            )
            vals = (bw.get_line1(twips), bw.get_line2(twips), bw.get_gap(twips))

        elif line == BorderLineStyleEnum.THINTHICK_MEDIUMGAP:
            flags = (
                mBwi.BorderWidthImplFlags.CHANGE_DIST
                | mBwi.BorderWidthImplFlags.CHANGE_LINE1
                | mBwi.BorderWidthImplFlags.CHANGE_LINE2
            )
            bw = mBwi.BorderWidthImpl(nFlags=flags, nRate1=0.5, nRate2=0.25, nRateGap=0.25)
            vals = (bw.get_line1(twips), bw.get_line2(twips), bw.get_gap(twips))

        elif line == BorderLineStyleEnum.THINTHICK_LARGEGAP:
            flags = mBwi.BorderWidthImplFlags.CHANGE_DIST
            bw = mBwi.BorderWidthImpl(
                nFlags=flags, nRate1=mBwi.THINTHICK_LARGEGAP_LINE1, nRate2=mBwi.THINTHICK_LARGEGAP_LINE2, nRateGap=1.0
            )
            vals = (bw.get_line1(twips), bw.get_line2(twips), bw.get_gap(twips))

        elif line == BorderLineStyleEnum.THICKTHIN_SMALLGAP:
            flags = mBwi.BorderWidthImplFlags.CHANGE_LINE2
            bw = mBwi.BorderWidthImpl(
                nFlags=flags, nRate1=mBwi.THICKTHIN_SMALLGAP_LINE1, nRate2=1.0, nRateGap=mBwi.THICKTHIN_SMALLGAP_GAP
            )
            vals = (bw.get_line1(twips), bw.get_line2(twips), bw.get_gap(twips))

        elif line == BorderLineStyleEnum.THICKTHIN_MEDIUMGAP:
            flags = (
                mBwi.BorderWidthImplFlags.CHANGE_DIST
                | mBwi.BorderWidthImplFlags.CHANGE_LINE1
                | mBwi.BorderWidthImplFlags.CHANGE_LINE2
            )
            bw = mBwi.BorderWidthImpl(nFlags=flags, nRate1=0.25, nRate2=0.5, nRateGap=0.25)
            vals = (bw.get_line1(twips), bw.get_line2(twips), bw.get_gap(twips))

        elif line == BorderLineStyleEnum.THICKTHIN_LARGEGAP:
            flags = mBwi.BorderWidthImplFlags.CHANGE_DIST
            bw = mBwi.BorderWidthImpl(
                nFlags=flags, nRate1=mBwi.THICKTHIN_LARGEGAP_LINE1, nRate2=mBwi.THICKTHIN_LARGEGAP_LINE2, nRateGap=1.0
            )
            vals = (bw.get_line1(twips), bw.get_line2(twips), bw.get_gap(twips))

        elif line in en_em:
            # Word compat: the lines widths are exactly following this rule, should be:
            # 0.75pt up to 3pt and then 3pt
            flags = (
                mBwi.BorderWidthImplFlags.CHANGE_DIST
                | mBwi.BorderWidthImplFlags.CHANGE_LINE1
                | mBwi.BorderWidthImplFlags.CHANGE_LINE2
            )
            bw = mBwi.BorderWidthImpl(nFlags=flags, nRate1=0.25, nRate2=0.25, nRateGap=0.5)
            vals = (bw.get_line1(twips), bw.get_line2(twips), bw.get_gap(twips))

        elif line == BorderLineStyleEnum.OUTSET:
            flags = flags = mBwi.BorderWidthImplFlags.CHANGE_DIST | mBwi.BorderWidthImplFlags.CHANGE_LINE2
            bw = mBwi.BorderWidthImpl(nFlags=flags, nRate1=mBwi.OUTSET_LINE1, nRate2=0.5, nRateGap=0.5)
            vals = (bw.get_line1(twips), bw.get_line2(twips), bw.get_gap(twips))

        elif line == BorderLineStyleEnum.INSET:
            flags = flags = mBwi.BorderWidthImplFlags.CHANGE_DIST | mBwi.BorderWidthImplFlags.CHANGE_LINE1
            bw = mBwi.BorderWidthImpl(nFlags=flags, nRate1=0.5, nRate2=mBwi.INSET_LINE2, nRateGap=0.5)
            vals = (bw.get_line1(twips), bw.get_line2(twips), bw.get_gap(twips))

        if vals:
            for key, value in zip(val_keys, vals):
                val = round(UnitConvert.convert_twip_mm100(value))
                self._set(key, val)

    # endregion init

    # region methods
    def __eq__(self, other: object) -> bool:
        bl2: BorderLine2 = None
        if isinstance(other, Side):
            bl2 = other.get_border_line2()
        elif getattr(other, "typeName", None) == "com.sun.star.table.BorderLine2":
            bl2 = other
        if bl2:
            bl1 = self.get_border_line2()
            return (
                bl1.Color == bl2.Color
                and bl1.InnerLineWidth == bl2.InnerLineWidth
                and bl1.LineDistance == bl2.LineDistance
                and bl1.LineStyle == bl2.LineStyle
                and bl1.LineWidth == bl2.LineWidth
                and bl1.OuterLineWidth == bl2.OuterLineWidth
            )
        return False

    def _supported_services(self) -> Tuple[str, ...]:
        return ()

    # region apply()
    @overload
    def apply(self, obj: object, *, flags: SideFlags) -> None:
        ...

    @overload
    def apply(self, obj: object, *, flags: SideFlags, keys: Dict[str, str]) -> None:
        ...

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies style to object

        Args:
            obj (object): Object to apply style to.

        Other Parameters:
            flags: (SideFlags): Determins where to apply side.
            keys: (Dict[str, str], optional): key map for properties.
                Can be any or all of the following ``left``, ``right``, ``top``, ``bottom``, ``diagonal_up``, ``diagonal_down``

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_APPLYING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_APPLYED` :eventref:`src-docs-event`

        Returns:
            None:
        """
        cargs = CancelEventArgs(source=f"{self.apply.__qualname__}")
        cargs.event_data = self
        self.on_applying(cargs)
        if cargs.cancel:
            return
        _Events().trigger(FormatNamedEvent.STYLE_APPLYING, cargs)
        if cargs.cancel:
            return

        expected_key_names = ("flags",)
        for kw in expected_key_names:
            if not kw in kwargs:
                raise Exception(f'apply_style() requires argument "{kw}"')
        flags = cast(SideFlags, kwargs["flags"])
        keys = {
            "left": "LeftBorder2",
            "right": "RightBorder2",
            "top": "TopBorder2",
            "bottom": "BottomBorder2",
            "diagonal_up": "DiagonalBLTR2",
            "diagonal_down": "DiagonalTLBR2",
        }
        if "keys" in kwargs:
            keys.update(kwargs["keys"])
        val = self.get_border_line2()
        applied = False
        if SideFlags.LEFT in flags:
            mProps.Props.set(obj, **{keys["left"]: val})
            applied == True
        if SideFlags.TOP in flags:
            mProps.Props.set(obj, **{keys["top"]: val})
            applied == True
        if SideFlags.RIGHT in flags:
            mProps.Props.set(obj, **{keys["right"]: val})
            applied == True
        if SideFlags.BOTTOM in flags:
            mProps.Props.set(obj, **{keys["bottom"]: val})
            applied == True
        if SideFlags.BOTTOM_LEFT_TOP_RIGHT in flags:
            mProps.Props.set(obj, **{keys["diagonal_up"]: val})
            applied == True
        if SideFlags.TOP_LEFT_BOTTOM_RIGHT in flags:
            mProps.Props.set(obj, **{keys["diagonal_down"]: val})
            applied == True

        if applied:
            eargs = EventArgs.from_args(cargs)
            self.on_applied(eargs)
            _Events().trigger(FormatNamedEvent.STYLE_APPLIED, eargs)

    # endregion apply()

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

    def copy(self) -> Side:
        """Gets a copy of current instance"""
        cp = super().copy()
        cp._pts = self._pts
        return cp

    @static_prop
    def empty() -> Side:
        """Gets an empyty side. When applied formatting is removed"""
        if Side._EMPTY is None:
            Side._EMPTY = Side(line=BorderLineStyleEnum.NONE, color=0, width=0.0)
        return Side._EMPTY

    @staticmethod
    def from_border2(border: BorderLine2) -> Side:
        side = Side()
        side._set("Color", border.Color)
        side._set("InnerLineWidth", border.InnerLineWidth)
        side._set("LineDistance", border.LineDistance)
        side._set("LineStyle", border.LineStyle)
        side._set("LineWidth", border.LineWidth)
        side._set("OuterLineWidth", border.OuterLineWidth)
        return side

    # endregion methods

    # region style methods
    def fmt_style(self, value: BorderLineStyleEnum) -> Side:
        """
        Gets copy of instance with style set.

        Args:
            value (BorderLineStyleEnum): style value

        Returns:
            Side: Side with style set
        """
        inst = Side(line=value, width=self._pts, color=self.prop_color)
        return inst

    def fmt_color(self, value: Color) -> Side:
        """
        Gets copy of instance with color set.

        Args:
            value (Color): color value

        Returns:
            Side: Side with color set
        """
        cp = self.copy()
        cp.prop_color = value
        return cp

    def fmt_width(self, value: float) -> Side:
        """
        Gets copy of instance with width set.

        Args:
            value (float): width value

        Returns:
            Side: Side with width set
        """
        inst = Side(line=self.prop_line, width=value, color=self.prop_color)
        return inst

    # endregion style methods
    # region Style Properties
    @property
    def line_none(self) -> Side:
        """Gets instance with no border line"""
        return self.fmt_style(BorderLineStyleEnum.NONE)

    @property
    def line_solid(self) -> Side:
        """Gets instance with solid border line"""
        return self.fmt_style(BorderLineStyleEnum.SOLID)

    @property
    def line_dotted(self) -> Side:
        """Gets instance with dotted border line"""
        return self.fmt_style(BorderLineStyleEnum.DOTTED)

    @property
    def line_dashed(self) -> Side:
        """Gets instance with dashed border line"""
        return self.fmt_style(BorderLineStyleEnum.DASHED)

    @property
    def line_dashed(self) -> Side:
        """Gets instance with dashed border line"""
        return self.fmt_style(BorderLineStyleEnum.DOUBLE)

    @property
    def line_thin_thick_small_gap(self) -> Side:
        """Gets instance with double border line with a thin line outside and a thick line inside separated by a small gap."""
        return self.fmt_style(BorderLineStyleEnum.THINTHICK_SMALLGAP)

    @property
    def line_thin_thick_medium_gap(self) -> Side:
        """Gets instance with double border line with a thin line outside and a thick line inside separated by a medium gap."""
        return self.fmt_style(BorderLineStyleEnum.THINTHICK_MEDIUMGAP)

    @property
    def line_thin_thick_large_gap(self) -> Side:
        """Gets instance with double border line with a thin line outside and a thick line inside separated by a large gap."""
        return self.fmt_style(BorderLineStyleEnum.THINTHICK_LARGEGAP)

    @property
    def line_thick_thin_small_gap(self) -> Side:
        """Gets instance with double border line with a thick line outside and a thin line inside separated by a small gap."""
        return self.fmt_style(BorderLineStyleEnum.THICKTHIN_SMALLGAP)

    @property
    def line_thick_thin_medium_gap(self) -> Side:
        """Gets instance with double border line with a thick line outside and a thin line inside separated by a medium gap."""
        return self.fmt_style(BorderLineStyleEnum.THICKTHIN_MEDIUMGAP)

    @property
    def line_thick_thin_large_gap(self) -> Side:
        """Gets instance with double border line with a thick line outside and a thin line inside separated by a large gap."""
        return self.fmt_style(BorderLineStyleEnum.THICKTHIN_LARGEGAP)

    @property
    def line_embossed(self) -> Side:
        """Gets instance with 3D embossed border line."""
        return self.fmt_style(BorderLineStyleEnum.EMBOSSED)

    @property
    def line_engraved(self) -> Side:
        """Gets instance with 3D engraved border line."""
        return self.fmt_style(BorderLineStyleEnum.ENGRAVED)

    @property
    def line_outset(self) -> Side:
        """Gets instance with outset border line."""
        return self.fmt_style(BorderLineStyleEnum.OUTSET)

    @property
    def line_inset(self) -> Side:
        """Gets instance with inset border line."""
        return self.fmt_style(BorderLineStyleEnum.INSET)

    @property
    def line_fine_dashed(self) -> Side:
        """Gets instance with finely dashed border line."""
        return self.fmt_style(BorderLineStyleEnum.FINE_DASHED)

    @property
    def line_double_thin(self) -> Side:
        """Gets instance with Double border line consisting of two fixed thin lines separated by a variable gap."""
        return self.fmt_style(BorderLineStyleEnum.DOUBLE_THIN)

    @property
    def line_dash_dot(self) -> Side:
        """Gets instance with line consisting of a repetition of one dash and one dot."""
        return self.fmt_style(BorderLineStyleEnum.DASH_DOT)

    @property
    def line_dash_dot_dot(self) -> Side:
        """Gets instance with line consisting of a repetition of one dash and 2 dots."""
        return self.fmt_style(BorderLineStyleEnum.DASH_DOT_DOT)

    # endregion Style Properties
    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.STRUCT

    @property
    def prop_line(self) -> BorderLineStyleEnum:
        """Gets Border Line style"""
        pv = cast(int, self._get("LineStyle"))
        return BorderLineStyleEnum(pv)

    @prop_line.setter
    def prop_line(self, value: BorderLineStyleEnum) -> None:
        self._set("LineStyle", value.value)
        self._set_line_values(self._pts, value)

    @property
    def prop_color(self) -> Color:
        """Gets Border Line Color"""
        return self._get("Color")

    @prop_color.setter
    def prop_color(self, value: Color) -> None:
        self._set("Color", value)

    @property
    def prop_width(self) -> float:
        """
        Gets Border Line Width.

        Contains the width of a single line or the width of outer part of a double line (in mm units).
        If this value is zero, no line is drawn.
        """
        return self._pts

    # endregion properties
