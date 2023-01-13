"""
Module for table side (``BorderLine2``) struct.

.. versionadded:: 0.9.0
"""
# region imports
from __future__ import annotations
from typing import Dict, Tuple, cast, overload
from enum import IntFlag

from ....meta.static_prop import static_prop
from ....utils import props as mProps
from ....utils.color import Color
from ....utils.color import CommonColor
from ...kind.style_kind import StyleKind
from ...style_base import StyleBase
from ...style_const import POINT_RATIO

import uno
from ooo.dyn.table.border_line import BorderLine as BorderLine
from ooo.dyn.table.border_line_style import BorderLineStyleEnum as BorderLineStyleEnum
from ooo.dyn.table.border_line2 import BorderLine2 as BorderLine2


# endregion imports

# region Enums


class SideFlags(IntFlag):
    """
    Side Flags Enum

    Any properties starting with ``prop_`` set or get current instance values.

    All methods ``style_`` can be used to chain together font properties.

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
        width: float = 0.75,
        width_inner: float = 0.0,
        distance: float = 0.0,
    ) -> None:
        """
        Constructs Side

        Args:
            line (BorderLineStyleEnum, optional): Line Style of the border.
            color (Color, optional): Color of the border.
            width (float, optional): Contains the width in of a single line or the width of outer part of a double line (in pt units). If this value is zero, no line is drawn. Default ``0.75``
            width_inner (float, optional): contains the width of the inner part of a double line (in pt units). If this value is zero, only a single line is drawn. Default ``0.0``
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
            "InnerLineWidth": round(width_inner * POINT_RATIO),
            "LineDistance": round(distance * 100),
            "LineStyle": line.value,
            "LineWidth": round(width * POINT_RATIO),
            "OuterLineWidth": round(width * POINT_RATIO),
        }

        super().__init__(**init_vals)

    # endregion init

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        return ()

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

        Returns:
            None:
        """
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
        if SideFlags.LEFT in flags:
            mProps.Props.set(obj, **{keys["left"]: val})
        if SideFlags.TOP in flags:
            mProps.Props.set(obj, **{keys["top"]: val})
        if SideFlags.RIGHT in flags:
            mProps.Props.set(obj, **{keys["right"]: val})
        if SideFlags.BOTTOM in flags:
            mProps.Props.set(obj, **{keys["bottom"]: val})
        if SideFlags.BOTTOM_LEFT_TOP_RIGHT in flags:
            mProps.Props.set(obj, **{keys["diagonal_up"]: val})
        if SideFlags.TOP_LEFT_BOTTOM_RIGHT in flags:
            mProps.Props.set(obj, **{keys["diagonal_down"]: val})

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
    def empty() -> Side:
        """Gets an empyty side. When applied formatting is removed"""
        if Side._EMPTY is None:
            Side._EMPTY = Side(line=BorderLineStyleEnum.NONE, color=0, width=0.0, width_inner=0.0, distance=0)
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

    # endregion methods

    # region style methods
    def style_style(self, value: BorderLineStyleEnum) -> Side:
        """
        Gets copy of instance with style set.

        Args:
            value (BorderLineStyleEnum): style value

        Returns:
            Side: Side with style set
        """
        cp = self.copy()
        cp.prop_line = value
        return cp

    def style_color(self, value: Color) -> Side:
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

    def style_width(self, value: float) -> Side:
        """
        Gets copy of instance with width set.

        Args:
            value (float): width value

        Returns:
            Side: Side with width set
        """
        cp = self.copy()
        cp.prop_width = value
        return cp

    def style_width_inner(self, value: float) -> Side:
        """
        Gets copy of instance with inner width set.

        Args:
            value (float): inner width value

        Returns:
            Side: Side with inner width set
        """
        cp = self.copy()
        cp.prop_width_inner = value
        return cp

    def style_distance(self, value: float) -> Side:
        """
        Gets copy of instance with distance set.

        Args:
            value (float): distance value

        Returns:
            Side: Side with distance set
        """
        cp = self.copy()
        cp.prop_distance = value
        return cp

    # endregion style methods
    # region Style Properties
    @property
    def line_none(self) -> Side:
        """Gets instance with no border line"""
        cp = self.copy()
        cp.prop_line = BorderLineStyleEnum.NONE
        return cp

    @property
    def line_solid(self) -> Side:
        """Gets instance with solid border line"""
        cp = self.copy()
        cp.prop_line = BorderLineStyleEnum.SOLID
        return cp

    @property
    def line_dotted(self) -> Side:
        """Gets instance with dotted border line"""
        cp = self.copy()
        cp.prop_line = BorderLineStyleEnum.DOTTED
        return cp

    @property
    def line_dashed(self) -> Side:
        """Gets instance with dashed border line"""
        cp = self.copy()
        cp.prop_line = BorderLineStyleEnum.DASHED
        return cp

    @property
    def line_dashed(self) -> Side:
        """Gets instance with dashed border line"""
        cp = self.copy()
        cp.prop_line = BorderLineStyleEnum.DOUBLE
        return cp

    @property
    def line_thin_thick_small_gap(self) -> Side:
        """Gets instance with double border line with a thin line outside and a thick line inside separated by a small gap."""
        cp = self.copy()
        cp.prop_line = BorderLineStyleEnum.THINTHICK_SMALLGAP
        return cp

    @property
    def line_thin_thick_medium_gap(self) -> Side:
        """Gets instance with double border line with a thin line outside and a thick line inside separated by a medium gap."""
        cp = self.copy()
        cp.prop_line = BorderLineStyleEnum.THINTHICK_MEDIUMGAP
        return cp

    @property
    def line_thin_thick_large_gap(self) -> Side:
        """Gets instance with double border line with a thin line outside and a thick line inside separated by a large gap."""
        cp = self.copy()
        cp.prop_line = BorderLineStyleEnum.THINTHICK_LARGEGAP
        return cp

    @property
    def line_thick_thin_small_gap(self) -> Side:
        """Gets instance with double border line with a thick line outside and a thin line inside separated by a small gap."""
        cp = self.copy()
        cp.prop_line = BorderLineStyleEnum.THICKTHIN_SMALLGAP
        return cp

    @property
    def line_thick_thin_medium_gap(self) -> Side:
        """Gets instance with double border line with a thick line outside and a thin line inside separated by a medium gap."""
        cp = self.copy()
        cp.prop_line = BorderLineStyleEnum.THICKTHIN_MEDIUMGAP
        return cp

    @property
    def line_thick_thin_large_gap(self) -> Side:
        """Gets instance with double border line with a thick line outside and a thin line inside separated by a large gap."""
        cp = self.copy()
        cp.prop_line = BorderLineStyleEnum.THICKTHIN_LARGEGAP
        return cp

    @property
    def line_embossed(self) -> Side:
        """Gets instance with 3D embossed border line."""
        cp = self.copy()
        cp.prop_line = BorderLineStyleEnum.EMBOSSED
        return cp

    @property
    def line_engraved(self) -> Side:
        """Gets instance with 3D engraved border line."""
        cp = self.copy()
        cp.prop_line = BorderLineStyleEnum.ENGRAVED
        return cp

    @property
    def line_outset(self) -> Side:
        """Gets instance with outset border line."""
        cp = self.copy()
        cp.prop_line = BorderLineStyleEnum.OUTSET
        return cp

    @property
    def line_inset(self) -> Side:
        """Gets instance with inset border line."""
        cp = self.copy()
        cp.prop_line = BorderLineStyleEnum.INSET
        return cp

    @property
    def line_fine_dashed(self) -> Side:
        """Gets instance with finely dashed border line."""
        cp = self.copy()
        cp.prop_line = BorderLineStyleEnum.FINE_DASHED
        return cp

    @property
    def line_double_thin(self) -> Side:
        """Gets instance with Double border line consisting of two fixed thin lines separated by a variable gap."""
        cp = self.copy()
        cp.prop_line = BorderLineStyleEnum.DOUBLE_THIN
        return cp

    @property
    def line_dash_dot(self) -> Side:
        """Gets instance with line consisting of a repetition of one dash and one dot."""
        cp = self.copy()
        cp.prop_line = BorderLineStyleEnum.DASH_DOT
        return cp

    @property
    def line_dash_dot_dot(self) -> Side:
        """Gets instance with line consisting of a repetition of one dash and 2 dots."""
        cp = self.copy()
        cp.prop_line = BorderLineStyleEnum.DASH_DOT_DOT
        return cp

    @property
    def line_border_line_style_max(self) -> Side:
        """Gets instance with maximum valid border line style value."""
        cp = self.copy()
        cp.prop_line = BorderLineStyleEnum.BORDER_LINE_STYLE_MAX
        return cp

    # endregion Style Properties
    # region properties
    @property
    def prop_style_kind(self) -> StyleKind:
        """Gets the kind of style"""
        return StyleKind.STRUCT

    @property
    def prop_line(self) -> BorderLineStyleEnum:
        """Gets Border Line style"""
        pv = cast(int, self._get("LineStyle"))
        return BorderLineStyleEnum(pv)

    @prop_line.setter
    def prop_line(self, value: BorderLineStyleEnum) -> None:
        self._set("LineStyle", value.value)

    @property
    def prop_color(self) -> Color:
        """Gets Border Line Color"""
        return self._get("Color")

    @prop_color.setter
    def prop_color(self, value: Color) -> None:
        self._set("Color", value)

    @property
    def prop_distance(self) -> float:
        """
        Gets/Sets the distance between the inner and outer parts of a double line (in mm units). Defalut ``0.0``
        """
        pv = cast(int, self._get("LineDistance"))
        if pv == 0:
            return 0.0
        return float(pv / 100)

    @prop_distance.setter
    def prop_distance(self, value: float):
        self._set("LineDistance", round(value, *100))

    @property
    def prop_width(self) -> float:
        """
        Gets Border Line Width.

        Contains the width of a single line or the width of outer part of a double line (in mm units).
        If this value is zero, no line is drawn.
        """
        pv = cast(int, self._get("LineWidth"))
        if pv == 0:
            return 0.0
        return float(pv / POINT_RATIO)

    @prop_width.setter
    def prop_width(self, value: float) -> None:
        i = round(value, *POINT_RATIO)
        self._set("LineWidth", i)
        self._set("OuterLineWidth", i)

    @property
    def prop_width_inner(self) -> float:
        """
        Gets Border Line Inner Width.

        Contains the width of the inner part of a double line (in mm units).
        If this value is zero, no line is drawn.
        """
        pv = cast(int, self._get("InnerLineWidth"))
        if pv == 0:
            return 0.0
        return float(pv / POINT_RATIO)

    @prop_width_inner.setter
    def prop_width_inner(self, value: float) -> None:
        self._set("InnerLineWidth", round(value, *POINT_RATIO))

    # endregion properties
