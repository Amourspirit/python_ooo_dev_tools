"""
Module for table side (``BorderLine2``) struct.

.. versionadded:: 0.9.0
"""

# region Import
from __future__ import annotations
import contextlib
from typing import Any, Tuple, cast, overload, Type, TypeVar, TYPE_CHECKING
from enum import Enum, IntEnum

import uno
from ooo.dyn.table.border_line import BorderLine
from ooo.dyn.table.border_line_style import BorderLineStyle
from ooo.dyn.table.border_line2 import BorderLine2

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.format.inner.common import border_width_impl as mBwi
from ooodev.format.inner.direct.structs.struct_base import StructBase
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.meta.deleted_enum_meta import DeletedUnoConstEnumMeta
from ooodev.mock import mock_g
from ooodev.units.unit_convert import UnitConvert
from ooodev.units.unit_convert import UnitLength
from ooodev.units.unit_obj import UnitT
from ooodev.units.unit_pt import UnitPT
from ooodev.utils import props as mProps
from ooodev.utils.color import Color
from ooodev.utils.color import StandardColor

# endregion Import

_TSide = TypeVar("_TSide", bound="Side")

# region Enums

if TYPE_CHECKING or mock_g.DOCS_BUILDING:
    # if doc are building then use this class for doc purposes.
    # Otherwise, Sphinx will error for BorderLineKind

    class BorderLineKind(IntEnum):
        """
        Enum of Const Class BorderLineStyle

        """

        NONE = BorderLineStyle.NONE
        """
        No border line.
        """
        SOLID = BorderLineStyle.SOLID
        """
        Solid border line.
        """
        DOTTED = BorderLineStyle.DOTTED
        """
        Dotted border line.
        """
        DASHED = BorderLineStyle.DASHED
        """
        Dashed border line.
        """
        DOUBLE = BorderLineStyle.DOUBLE
        """
        Double border line.
        
        Widths of the lines and the gap are all equal, and vary equally with the total width.
        """
        THINTHICK_SMALLGAP = BorderLineStyle.THINTHICK_SMALLGAP
        """
        Double border line with a thin line outside and a thick line inside separated by a small gap.
        """
        THINTHICK_MEDIUMGAP = BorderLineStyle.THINTHICK_MEDIUMGAP
        """
        Double border line with a thin line outside and a thick line inside separated by a medium gap.
        """
        THINTHICK_LARGEGAP = BorderLineStyle.THINTHICK_LARGEGAP
        """
        Double border line with a thin line outside and a thick line inside separated by a large gap.
        """
        THICKTHIN_SMALLGAP = BorderLineStyle.THICKTHIN_SMALLGAP
        """
        Double border line with a thick line outside and a thin line inside separated by a small gap.
        """
        THICKTHIN_MEDIUMGAP = BorderLineStyle.THICKTHIN_MEDIUMGAP
        """
        Double border line with a thick line outside and a thin line inside separated by a medium gap.
        """
        THICKTHIN_LARGEGAP = BorderLineStyle.THICKTHIN_LARGEGAP
        """
        Double border line with a thick line outside and a thin line inside separated by a large gap.
        """
        EMBOSSED = BorderLineStyle.EMBOSSED
        """
        3D embossed border line.
        """
        ENGRAVED = BorderLineStyle.ENGRAVED
        """
        3D engraved border line.
        """
        OUTSET = BorderLineStyle.OUTSET
        """
        Outset border line.
        """
        INSET = BorderLineStyle.INSET
        """
        Inset border line.
        """
        FINE_DASHED = BorderLineStyle.FINE_DASHED
        """
        Finely dashed border line.
        """
        DOUBLE_THIN = BorderLineStyle.DOUBLE_THIN
        """
        Double border line consisting of two fixed thin lines separated by a variable gap.
        """
        DASH_DOT = BorderLineStyle.DASH_DOT
        """
        Line consisting of a repetition of one dash and one dot.
        """
        DASH_DOT_DOT = BorderLineStyle.DASH_DOT_DOT
        """
        Line consisting of a repetition of one dash and 2 dots.
        """
        BORDER_LINE_STYLE_MAX = BorderLineStyle.BORDER_LINE_STYLE_MAX
        """
        Maximum valid border line style value.
        """

else:
    # Class takes the place of the above class at runtime.
    # The reason for this to make sure 'AT_FRAME' enum value is excluded.
    # Also, future-proof enum, if later version add new enum values.
    class BorderLineKind(
        IntEnum,
        metaclass=DeletedUnoConstEnumMeta,
        type_name="com.sun.star.table.BorderLineStyle",
        name_space="com.sun.star.table",
    ):
        """Dynamic Enum. Contains all the constant values of ``com.sun.star.table.BorderLineStyle`` as Enum values"""

        @staticmethod
        def _get_deleted_attribs() -> Tuple[str]:
            return ("BORDER_LINE_STYLE_MAX",)


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
    """``2.25pt``"""
    EXTRA_THICK = (5, 4.5)
    """``4.5pt``"""

    def __float__(self) -> float:
        return self.value[1]

    def get_value_pt(self) -> float:
        """
        Gets instance value in ``pt`` (point) units.

        Returns:
            float: Value in ``pt`` units.
        """
        return self.value[1]


# endregion Enums

# from some reason LibreOffice sometimes changes the values of BorderLine2 value with certain BorderLineStyleEnum
# such as DOUBLE_THIN, in testing is showed that DOUBLE_THIN caused BorderLine2.LineWidth to be changed.
# However, setting DOUBLE_THIN in libreOffice manually does not have this effect.


class Side(StructBase):
    """
    Side struct.

    Represents one side of a border.

    .. versionadded:: 0.9.0
    """

    # region init

    def __init__(
        self,
        *,
        line: BorderLineKind = BorderLineKind.SOLID,
        color: Color = StandardColor.BLACK,
        width: LineSize | float | UnitT = LineSize.THIN,
    ) -> None:
        """
        Constructs Side

        Args:
            line (BorderLineStyleEnum, optional): Line Style of the border. Default ``BorderLineKind.SOLID``.
            color (:py:data:`~.utils.color.Color`, optional): Color of the border. Default ``StandardColor.BLACK``
            width (LineSize, float, UnitT, optional): Contains the width in of a single line or the width of outer part of a double line (in ``pt`` units) or :ref:`proto_unit_obj`. If this value is zero, no line is drawn. Default ``LineSize.THIN``

        Raises:
            ValueError: if ``color``, ``width`` or ``width_inner`` is less than ``0``.

        Returns:
            None:
        """
        try:
            self._pts = cast(float, width.get_value_pt())  # type: ignore
        except AttributeError:
            self._pts = float(width)  # type: ignore

        if color < 0:
            raise ValueError("color must be a positive value")
        if self._pts < 0.0:
            raise ValueError("width must be a positive value")
        if self._pts > 9.0000001:
            raise ValueError("Maximum width allowed is 9pt")

        lw = round(UnitConvert.convert(num=self._pts, frm=UnitLength.PT, to=UnitLength.MM100))

        init_vals = {
            "Color": color,
            "InnerLineWidth": 0,
            "LineDistance": 0,
            "LineStyle": line.value,
            "LineWidth": lw,
            "OuterLineWidth": 0,
        }

        super().__init__(**init_vals)
        self._set_line_values(pts=self._pts, line=line)

    def _set_line_values(self, pts: float, line: BorderLineKind) -> None:
        # sourcery skip: extract-duplicate-method, inline-immediately-returned-variable, low-code-quality
        # taken care of by creating dynamic enum BorderLineKind
        # if line.name == BorderLineKind.BORDER_LINE_STYLE_MAX:
        #     raise ValueError("BORDER_LINE_STYLE_MAX is not supported")

        val_keys = ("OuterLineWidth", "InnerLineWidth", "LineDistance")

        if line == BorderLineKind.NONE:
            for attr in val_keys:
                self._set(attr, 0)
            return

        twips = round(UnitConvert.to_twips(pts, UnitLength.PT))
        single_lns = (
            BorderLineKind.SOLID,
            BorderLineKind.DOTTED,
            BorderLineKind.DASHED,
            BorderLineKind.FINE_DASHED,
            BorderLineKind.DASH_DOT,
            BorderLineKind.DASH_DOT_DOT,
        )
        en_em = (BorderLineKind.ENGRAVED, BorderLineKind.EMBOSSED)

        vals = None

        if line in single_lns:
            flags = mBwi.BorderWidthImplFlags.CHANGE_LINE1
            bw = mBwi.BorderWidthImpl(nFlags=flags, nRate1=1.0, nRate2=0.0, nRateGap=0.0)
            vals = (bw.get_line1(twips), bw.get_line2(twips), bw.get_gap(twips))

        elif line == BorderLineKind.DOUBLE:
            flags = (
                mBwi.BorderWidthImplFlags.CHANGE_DIST
                | mBwi.BorderWidthImplFlags.CHANGE_LINE1
                | mBwi.BorderWidthImplFlags.CHANGE_LINE2
            )
            bw = mBwi.BorderWidthImpl(nFlags=flags, nRate1=1 / 3, nRate2=1 / 3, nRateGap=1 / 3)
            vals = (bw.get_line1(twips), bw.get_line2(twips), bw.get_gap(twips))

        elif line == BorderLineKind.DOUBLE_THIN:
            flags = mBwi.BorderWidthImplFlags.CHANGE_DIST
            bw = mBwi.BorderWidthImpl(nFlags=flags, nRate1=10.0, nRate2=10.0, nRateGap=1.0)
            vals = (bw.get_line1(twips), bw.get_line2(twips), bw.get_gap(twips))

        elif line == BorderLineKind.THINTHICK_SMALLGAP:
            flags = mBwi.BorderWidthImplFlags.CHANGE_LINE1
            bw = mBwi.BorderWidthImpl(
                nFlags=flags, nRate1=1.0, nRate2=mBwi.THINTHICK_SMALLGAP_LINE2, nRateGap=mBwi.THINTHICK_SMALLGAP_GAP
            )
            vals = (bw.get_line1(twips), bw.get_line2(twips), bw.get_gap(twips))

        elif line == BorderLineKind.THINTHICK_MEDIUMGAP:
            flags = (
                mBwi.BorderWidthImplFlags.CHANGE_DIST
                | mBwi.BorderWidthImplFlags.CHANGE_LINE1
                | mBwi.BorderWidthImplFlags.CHANGE_LINE2
            )
            bw = mBwi.BorderWidthImpl(nFlags=flags, nRate1=0.5, nRate2=0.25, nRateGap=0.25)
            vals = (bw.get_line1(twips), bw.get_line2(twips), bw.get_gap(twips))

        elif line == BorderLineKind.THINTHICK_LARGEGAP:
            flags = mBwi.BorderWidthImplFlags.CHANGE_DIST
            bw = mBwi.BorderWidthImpl(
                nFlags=flags, nRate1=mBwi.THINTHICK_LARGEGAP_LINE1, nRate2=mBwi.THINTHICK_LARGEGAP_LINE2, nRateGap=1.0
            )
            vals = (bw.get_line1(twips), bw.get_line2(twips), bw.get_gap(twips))

        elif line == BorderLineKind.THICKTHIN_SMALLGAP:
            flags = mBwi.BorderWidthImplFlags.CHANGE_LINE2
            bw = mBwi.BorderWidthImpl(
                nFlags=flags, nRate1=mBwi.THICKTHIN_SMALLGAP_LINE1, nRate2=1.0, nRateGap=mBwi.THICKTHIN_SMALLGAP_GAP
            )
            vals = (bw.get_line1(twips), bw.get_line2(twips), bw.get_gap(twips))

        elif line == BorderLineKind.THICKTHIN_MEDIUMGAP:
            flags = (
                mBwi.BorderWidthImplFlags.CHANGE_DIST
                | mBwi.BorderWidthImplFlags.CHANGE_LINE1
                | mBwi.BorderWidthImplFlags.CHANGE_LINE2
            )
            bw = mBwi.BorderWidthImpl(nFlags=flags, nRate1=0.25, nRate2=0.5, nRateGap=0.25)
            vals = (bw.get_line1(twips), bw.get_line2(twips), bw.get_gap(twips))

        elif line == BorderLineKind.THICKTHIN_LARGEGAP:
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

        elif line == BorderLineKind.OUTSET:
            flags = flags = mBwi.BorderWidthImplFlags.CHANGE_DIST | mBwi.BorderWidthImplFlags.CHANGE_LINE2
            bw = mBwi.BorderWidthImpl(nFlags=flags, nRate1=mBwi.OUTSET_LINE1, nRate2=0.5, nRateGap=0.5)
            vals = (bw.get_line1(twips), bw.get_line2(twips), bw.get_gap(twips))

        elif line == BorderLineKind.INSET:
            flags = flags = mBwi.BorderWidthImplFlags.CHANGE_DIST | mBwi.BorderWidthImplFlags.CHANGE_LINE1
            bw = mBwi.BorderWidthImpl(nFlags=flags, nRate1=0.5, nRate2=mBwi.INSET_LINE2, nRateGap=0.5)
            vals = (bw.get_line1(twips), bw.get_line2(twips), bw.get_gap(twips))

        if vals:
            for key, value in zip(val_keys, vals):
                val = round(UnitConvert.convert_twip_mm100(value))
                self._set(key, val)

    # endregion init

    # region Overrides
    def _get_internal_cattribs(self) -> dict:
        cattribs = {
            "_props_internal_attributes": self._props,
            "_supported_services_values": self._supported_services(),
            "_format_kind_prop": self.prop_format_kind,
        }
        with contextlib.suppress(NotImplementedError):
            cattribs["_property_name"] = self._get_property_name()
        return cattribs

    def __eq__(self, other: Any) -> bool:
        # noinspection PyTypeChecker
        bl2 = None
        if isinstance(other, Side):
            bl2 = other.get_uno_struct()
        elif getattr(other, "typeName", None) == "com.sun.star.table.BorderLine2":
            bl2 = cast(BorderLine2, other)
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

    def _on_modifying(self, source: Any, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifying(source, event)

    # region apply()

    @overload
    def apply(self, obj: Any) -> None:  # type: ignore
        ...

    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies style to object

        Args:
            obj (object): Object to apply style to.

        Returns:
            None:
        """
        if not mProps.Props.has(obj, self._get_property_name()):
            self._print_not_valid_srv("apply")
            return

        struct = self.get_uno_struct()
        props = {self.property_name: struct}
        super().apply(obj=obj, override_dv=props)

    # endregion apply()

    # region copy()
    @overload
    def copy(self: _TSide) -> _TSide: ...

    @overload
    def copy(self: _TSide, **kwargs) -> _TSide: ...

    def copy(self: _TSide, **kwargs) -> _TSide:
        """Gets a copy of instance as a new instance"""
        # pylint: disable=protected-access
        cp = super().copy(**kwargs)
        cp._pts = self._pts
        self._copy_missing_attribs(self, cp, "_property_name", "_supported_services_values")
        return cp

    # endregion copy()

    # endregion Overrides

    # region methods

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
        line = BorderLine2()  # create the borderline
        for key, val in self._dv.items():
            setattr(line, key, val)
        return line

    # endregion methods

    # region static methods
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TSide], obj: Any) -> _TSide: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TSide], obj: Any, **kwargs) -> _TSide: ...

    @classmethod
    def from_obj(cls: Type[_TSide], obj: Any, **kwargs) -> _TSide:
        """
        Gets instance from object

        Args:
            obj (object): UNO object

        Returns:
            Side: Instance from object
        """

        nu = cls(**kwargs)

        border = cast(BorderLine2, mProps.Props.get(obj, nu._get_property_name()))
        return cls.from_uno_struct(border, **kwargs)

    # endregion from_obj()

    # region from_uno_struct()
    @overload
    @classmethod
    def from_uno_struct(cls: Type[_TSide], border: BorderLine2) -> _TSide: ...

    @overload
    @classmethod
    def from_uno_struct(cls: Type[_TSide], border: BorderLine2, **kwargs) -> _TSide: ...

    @classmethod
    def from_uno_struct(cls: Type[_TSide], border: BorderLine2, **kwargs) -> _TSide:
        """
        Gets instance that is populated from UNO ``BorderLine2``.

        Args:
            border (BorderLine2): UNO struct

        Returns:
            Side: instance.
        """
        # pylint: disable=protected-access
        pt_width = round(UnitConvert.convert(num=border.LineWidth, frm=UnitLength.MM100, to=UnitLength.PT), 2)
        inst = cls(width=pt_width, **kwargs)
        inst._set("Color", border.Color)
        inst._set("InnerLineWidth", border.InnerLineWidth)
        inst._set("LineDistance", border.LineDistance)
        inst._set("LineStyle", border.LineStyle)
        inst._set("LineWidth", border.LineWidth)
        inst._set("OuterLineWidth", border.OuterLineWidth)
        return inst

    # endregion from_uno_struct()

    # endregion static methods

    # region style methods
    def fmt_style(self: _TSide, value: BorderLineKind) -> _TSide:
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
        """Gets instance with no borderline"""
        return self.fmt_style(BorderLineKind.NONE)

    @property
    def line_solid(self: _TSide) -> _TSide:
        """Gets instance with solid borderline"""
        return self.fmt_style(BorderLineKind.SOLID)

    @property
    def line_dotted(self: _TSide) -> _TSide:
        """Gets instance with dotted borderline"""
        return self.fmt_style(BorderLineKind.DOTTED)

    @property
    def line_dashed(self: _TSide) -> _TSide:
        """Gets instance with dashed borderline"""
        return self.fmt_style(BorderLineKind.DASHED)

    @property
    def line_double(self: _TSide) -> _TSide:
        """Gets instance with double borderline"""
        return self.fmt_style(BorderLineKind.DOUBLE)

    @property
    def line_thin_thick_small_gap(self: _TSide) -> _TSide:
        """
        Gets instance with double borderline with a thin line outside and a thick line inside separated by a small gap.
        """
        return self.fmt_style(BorderLineKind.THINTHICK_SMALLGAP)

    @property
    def line_thin_thick_medium_gap(self: _TSide) -> _TSide:
        """
        Gets instance with double borderline with a thin line outside and a thick line inside separated by a medium gap.
        """
        return self.fmt_style(BorderLineKind.THINTHICK_MEDIUMGAP)

    @property
    def line_thin_thick_large_gap(self: _TSide) -> _TSide:
        """
        Gets instance with double borderline with a thin line outside and a thick line inside separated by a large gap.
        """
        return self.fmt_style(BorderLineKind.THINTHICK_LARGEGAP)

    @property
    def line_thick_thin_small_gap(self: _TSide) -> _TSide:
        """
        Gets instance with double borderline with a thick line outside and a thin line inside separated by a small gap.
        """
        return self.fmt_style(BorderLineKind.THICKTHIN_SMALLGAP)

    @property
    def line_thick_thin_medium_gap(self: _TSide) -> _TSide:
        """
        Gets instance with double borderline with a thick line outside and a thin line inside separated by a medium gap.
        """
        return self.fmt_style(BorderLineKind.THICKTHIN_MEDIUMGAP)

    @property
    def line_thick_thin_large_gap(self: _TSide) -> _TSide:
        """
        Gets instance with double borderline with a thick line outside and a thin line inside separated by a large gap.
        """
        return self.fmt_style(BorderLineKind.THICKTHIN_LARGEGAP)

    @property
    def line_embossed(self: _TSide) -> _TSide:
        """Gets instance with 3D embossed borderline."""
        return self.fmt_style(BorderLineKind.EMBOSSED)

    @property
    def line_engraved(self: _TSide) -> _TSide:
        """Gets instance with 3D engraved borderline."""
        return self.fmt_style(BorderLineKind.ENGRAVED)

    @property
    def line_outset(self: _TSide) -> _TSide:
        """Gets instance with outset borderline."""
        return self.fmt_style(BorderLineKind.OUTSET)

    @property
    def line_inset(self: _TSide) -> _TSide:
        """Gets instance with inset borderline."""
        return self.fmt_style(BorderLineKind.INSET)

    @property
    def line_fine_dashed(self: _TSide) -> _TSide:
        """Gets instance with finely dashed borderline."""
        return self.fmt_style(BorderLineKind.FINE_DASHED)

    @property
    def line_double_thin(self: _TSide) -> _TSide:
        """Gets instance with Double borderline consisting of two fixed thin lines separated by a variable gap."""
        return self.fmt_style(BorderLineKind.DOUBLE_THIN)

    @property
    def line_dash_dot(self: _TSide) -> _TSide:
        """Gets instance with line consisting of a repetition of one dash and one dot."""
        return self.fmt_style(BorderLineKind.DASH_DOT)

    @property
    def line_dash_dot_dot(self: _TSide) -> _TSide:
        """Gets instance with line consisting of a repetition of one dash and 2 dots."""
        return self.fmt_style(BorderLineKind.DASH_DOT_DOT)

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
    def prop_line(self) -> BorderLineKind:
        """Gets borderline style"""
        pv = cast(int, self._get("LineStyle"))
        return BorderLineKind(pv)

    @prop_line.setter
    def prop_line(self, value: BorderLineKind) -> None:
        self._set("LineStyle", value.value)
        self._set_line_values(self._pts, value)

    @property
    def prop_color(self) -> Color:
        """Gets borderline Color"""
        return self._get("Color")

    @prop_color.setter
    def prop_color(self, value: Color) -> None:
        self._set("Color", value)

    @property
    def property_name(self) -> str:
        """
        Gets/Sets property name

        This is the name of the property that the side should be applied to. Such as ``LeftBorder``, ``RightBorder`` etc.
        """
        return self._get_property_name()

    @property_name.setter
    def property_name(self, value: str) -> None:
        self._set_property_name(value)

    @property
    def prop_width(self) -> UnitPT:
        """
        Gets borderline Width.

        Contains the width of a single line or the width of outer part of a double line (in ``pt`` units).
        If this value is zero, no line is drawn.
        """
        return UnitPT(self._pts)

    @property
    def empty(self: _TSide) -> _TSide:
        """Gets an empty side. When applied formatting is removed"""
        # pylint: disable=protected-access
        # pylint: disable=unexpected-keyword-arg
        try:
            return self._empty_inst
        except AttributeError:
            self._empty_inst = self.__class__(
                line=BorderLineKind.NONE, color=0, width=0.0, _cattribs=self._get_internal_cattribs()  # type: ignore
            )
            self._empty_inst._is_default_inst = True
        return self._empty_inst

    # endregion properties
