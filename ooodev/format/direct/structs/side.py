"""
Module for table side (``BorderLine2``) struct.

.. versionadded:: 0.9.0
"""
# region imports
from __future__ import annotations
from typing import Dict, Tuple, cast, overload, Type, TypeVar
from enum import IntFlag, Enum

import uno
from ....events.event_singleton import _Events
from ....meta.static_prop import static_prop
from ....utils import props as mProps
from ....utils.color import Color
from ....utils.color import CommonColor
from ...kind.format_kind import FormatKind
from ...style_base import StyleBase, EventArgs, CancelEventArgs, FormatNamedEvent
from ..common import border_width_impl as mBwi
from ....utils.unit_convert import UnitConvert, Length

from ooo.dyn.table.border_line import BorderLine as BorderLine
from ooo.dyn.table.border_line_style import BorderLineStyleEnum as BorderLineStyleEnum
from ooo.dyn.table.border_line2 import BorderLine2 as BorderLine2


# endregion imports

_TSide = TypeVar(name="_TSide", bound="Side")

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

    # region init

    def __init__(
        self,
        *,
        line: BorderLineStyleEnum = BorderLineStyleEnum.SOLID,
        color: Color = CommonColor.BLACK,
        width: LineSize | float = LineSize.THIN,
    ) -> None:
        """
        Constructs Side

        Args:
            line (BorderLineStyleEnum, optional): Line Style of the border.
            color (Color, optional): Color of the border.
            width (LineSize, float, optional): Contains the width in of a single line or the width of outer part of a double line (in `pt` units). If this value is zero, no line is drawn. Default ``0.75``

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
            bl2 = other.get_uno_struct()
        elif getattr(other, "typeName", None) == "com.sun.star.table.BorderLine2":
            bl2 = other
        if bl2:
            bl1 = self.get_uno_struct()
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
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ()
        return self._supported_services_values

    def _on_modifing(self, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifing(event)

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
        val = self.get_uno_struct()
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

    def get_uno_struct_border_line(self) -> BorderLine:
        """
        Gets UNO ``BorderLine`` from instance.

        Returns:
            BorderLine: ``BorderLine`` instance
        """
        b2 = self.get_uno_struct()
        return BorderLine(
            Color=b2.Color,
            InnerLineWidth=b2.InnerLineWidth,
            OuterLineWidth=b2.OuterLineWidth,
            LineDistance=b2.LineDistance,
        )

    def get_uno_struct(self) -> BorderLine2:
        """
        Gets UNO ``BorderLine2`` from instance.

        Returns:
            BorderLine2: ``BorderLine2`` instance
        """
        line = BorderLine2()  # create the border line
        for key, val in self._dv.items():
            setattr(line, key, val)
        return line

    # region copy()
    @overload
    def copy(self: _TSide) -> _TSide:
        ...

    @overload
    def copy(self: _TSide, **kwargs) -> _TSide:
        ...

    def copy(self: _TSide, **kwargs) -> _TSide:
        """Gets a copy of instance as a new instance"""
        cp = super().copy(**kwargs)
        cp._pts = self._pts
        return cp

    # endregion copy()

    @static_prop
    def empty() -> Side:
        """Gets an empyty side. When applied formatting is removed"""
        try:
            return Side._EMPTY_INST
        except AttributeError:
            Side._EMPTY_INST = Side(line=BorderLineStyleEnum.NONE, color=0, width=0.0)
            Side._EMPTY_INST._is_default_inst = True
        return Side._EMPTY_INST

    # region from_border2()
    @overload
    @classmethod
    def from_uno_struct(cls: Type[_TSide], border: BorderLine2) -> _TSide:
        ...

    @overload
    @classmethod
    def from_uno_struct(cls: Type[_TSide], border: BorderLine2, **kwargs) -> _TSide:
        ...

    @classmethod
    def from_uno_struct(cls: Type[_TSide], border: BorderLine2, **kwargs) -> _TSide:
        """
        Gets instance that is populated from UNO ``BorderLine2``.

        Args:
            border (BorderLine2): UNO struct

        Returns:
            Side: instance.
        """
        pt_width = round(UnitConvert.convert(num=border.LineWidth, frm=Length.MM100, to=Length.PT), 2)
        inst = cls(width=pt_width, **kwargs)
        inst._set("Color", border.Color)
        inst._set("InnerLineWidth", border.InnerLineWidth)
        inst._set("LineDistance", border.LineDistance)
        inst._set("LineStyle", border.LineStyle)
        inst._set("LineWidth", border.LineWidth)
        inst._set("OuterLineWidth", border.OuterLineWidth)
        return inst

    # endregion from_border2()
    # endregion methods

    # region style methods
    def fmt_style(self: _TSide, value: BorderLineStyleEnum) -> _TSide:
        """
        Gets copy of instance with style set.

        Args:
            value (BorderLineStyleEnum): style value

        Returns:
            Side: Side with style set
        """
        return self.__class__(line=value, width=self._pts, color=self.prop_color)

    def fmt_color(self: _TSide, value: Color) -> _TSide:
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

    def fmt_width(self: _TSide, value: float) -> _TSide:
        """
        Gets copy of instance with width set.

        Args:
            value (float): width value

        Returns:
            Side: Side with width set
        """
        return self.__class__(line=self.prop_line, width=value, color=self.prop_color)

    # endregion style methods
    # region Style Properties
    @property
    def line_none(self: _TSide) -> _TSide:
        """Gets instance with no border line"""
        return self.fmt_style(BorderLineStyleEnum.NONE)

    @property
    def line_solid(self: _TSide) -> _TSide:
        """Gets instance with solid border line"""
        return self.fmt_style(BorderLineStyleEnum.SOLID)

    @property
    def line_dotted(self: _TSide) -> _TSide:
        """Gets instance with dotted border line"""
        return self.fmt_style(BorderLineStyleEnum.DOTTED)

    @property
    def line_dashed(self: _TSide) -> _TSide:
        """Gets instance with dashed border line"""
        return self.fmt_style(BorderLineStyleEnum.DASHED)

    @property
    def line_dashed(self: _TSide) -> _TSide:
        """Gets instance with dashed border line"""
        return self.fmt_style(BorderLineStyleEnum.DOUBLE)

    @property
    def line_thin_thick_small_gap(self: _TSide) -> _TSide:
        """Gets instance with double border line with a thin line outside and a thick line inside separated by a small gap."""
        return self.fmt_style(BorderLineStyleEnum.THINTHICK_SMALLGAP)

    @property
    def line_thin_thick_medium_gap(self: _TSide) -> _TSide:
        """Gets instance with double border line with a thin line outside and a thick line inside separated by a medium gap."""
        return self.fmt_style(BorderLineStyleEnum.THINTHICK_MEDIUMGAP)

    @property
    def line_thin_thick_large_gap(self: _TSide) -> _TSide:
        """Gets instance with double border line with a thin line outside and a thick line inside separated by a large gap."""
        return self.fmt_style(BorderLineStyleEnum.THINTHICK_LARGEGAP)

    @property
    def line_thick_thin_small_gap(self: _TSide) -> _TSide:
        """Gets instance with double border line with a thick line outside and a thin line inside separated by a small gap."""
        return self.fmt_style(BorderLineStyleEnum.THICKTHIN_SMALLGAP)

    @property
    def line_thick_thin_medium_gap(self: _TSide) -> _TSide:
        """Gets instance with double border line with a thick line outside and a thin line inside separated by a medium gap."""
        return self.fmt_style(BorderLineStyleEnum.THICKTHIN_MEDIUMGAP)

    @property
    def line_thick_thin_large_gap(self: _TSide) -> _TSide:
        """Gets instance with double border line with a thick line outside and a thin line inside separated by a large gap."""
        return self.fmt_style(BorderLineStyleEnum.THICKTHIN_LARGEGAP)

    @property
    def line_embossed(self: _TSide) -> _TSide:
        """Gets instance with 3D embossed border line."""
        return self.fmt_style(BorderLineStyleEnum.EMBOSSED)

    @property
    def line_engraved(self: _TSide) -> _TSide:
        """Gets instance with 3D engraved border line."""
        return self.fmt_style(BorderLineStyleEnum.ENGRAVED)

    @property
    def line_outset(self: _TSide) -> _TSide:
        """Gets instance with outset border line."""
        return self.fmt_style(BorderLineStyleEnum.OUTSET)

    @property
    def line_inset(self: _TSide) -> _TSide:
        """Gets instance with inset border line."""
        return self.fmt_style(BorderLineStyleEnum.INSET)

    @property
    def line_fine_dashed(self: _TSide) -> _TSide:
        """Gets instance with finely dashed border line."""
        return self.fmt_style(BorderLineStyleEnum.FINE_DASHED)

    @property
    def line_double_thin(self: _TSide) -> _TSide:
        """Gets instance with Double border line consisting of two fixed thin lines separated by a variable gap."""
        return self.fmt_style(BorderLineStyleEnum.DOUBLE_THIN)

    @property
    def line_dash_dot(self: _TSide) -> _TSide:
        """Gets instance with line consisting of a repetition of one dash and one dot."""
        return self.fmt_style(BorderLineStyleEnum.DASH_DOT)

    @property
    def line_dash_dot_dot(self: _TSide) -> _TSide:
        """Gets instance with line consisting of a repetition of one dash and 2 dots."""
        return self.fmt_style(BorderLineStyleEnum.DASH_DOT_DOT)

    # endregion Style Properties
    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.STRUCT
        return self._format_kind_prop

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
