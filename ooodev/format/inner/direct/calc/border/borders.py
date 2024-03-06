"""
Module for managing table borders (cells and ranges).

.. versionadded:: 0.9.0
"""

# region imports
from __future__ import annotations
from typing import Any, Type, overload, cast, Tuple, TypeVar

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.exceptions import ex as mEx
from ooodev.units.unit_obj import UnitT
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleMulti
from ooodev.format.inner.common.props.border_props import BorderProps as BorderProps
from ooodev.format.inner.common.props.cell_borders_props import CellBordersProps
from ooodev.format.inner.common.props.prop_pair import PropPair
from ooodev.format.inner.common.props.struct_border_table_props import StructBorderTableProps
from ooodev.format.inner.direct.structs.side import Side as Side
from ooodev.format.inner.direct.structs.table_border_struct import TableBorderStruct
from ooodev.format.inner.direct.calc.border.padding import Padding as Padding
from ooodev.format.inner.direct.calc.border.shadow import Shadow as Shadow


# endregion imports

_TBorders = TypeVar(name="_TBorders", bound="Borders")


class Borders(StyleMulti):
    """
    Table Borders used in styles for table cells and ranges.

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. seealso::

        - :ref:`help_calc_format_direct_cell_borders`

    .. versionadded:: 0.9.0
    """

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
        distance: float | UnitT | None = None,
        diagonal_down: Side | None = None,
        diagonal_up: Side | None = None,
        shadow: Shadow | None = None,
        padding: Padding | None = None,
    ) -> None:
        """
        Constructor

        Args:
            left (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style at the left edge.
            right (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style at the right edge.
            top (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style at the top edge.
            bottom (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style at the bottom edge.
            border_side (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style at the top, bottom, left, right edges. If this argument has a value then arguments ``top``, ``bottom``, ``left``, ``right`` are ignored
            horizontal (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style of horizontal lines for the inner part of a cell range.
            vertical (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style of vertical lines for the inner part of a cell range.
            distance (float, UnitT, optional): Contains the distance between the lines and other contents in ``mm`` units or :ref:`proto_unit_obj`.
            diagonal_down (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style from top-left to bottom-right diagonal.
            diagonal_up (~ooodev.format.inner.direct.structs.side.Side, optional): Specifies the line style from bottom-left to top-right diagonal.
            shadow (~ooodev.format.inner.direct.calc.border.shadow.Shadow, optional): Cell Shadow.
            padding (padding, optional): Cell padding.

        Returns:
            None:

        See Also:

            - :ref:`help_calc_format_direct_cell_borders`
        """
        # pylint: disable=unexpected-keyword-arg
        init_vals = {}

        if shadow is None:
            shadow_fmt = None
        else:
            shadow_fmt = shadow.copy(_cattribs=self._get_shadow_cattribs())

        if diagonal_down is None:
            diag_dn = None
        else:
            diag_dn = diagonal_down.copy(_cattribs=self._get_diagonal_dn_cattribs())

        if diagonal_up is None:
            diag_up = None
        else:
            diag_up = diagonal_up.copy(_cattribs=self._get_diagonal_up_cattribs())

        if padding is None:
            padding_fmt = None
        else:
            padding_fmt = padding.copy(_cattribs=self._get_padding_cattribs())

        border_table = TableBorderStruct(
            left=left,
            right=right,
            top=top,
            bottom=bottom,
            border_side=border_side,
            horizontal=horizontal,
            vertical=vertical,
            distance=distance,
            _cattribs=self._get_tb_cattribs(),  # type: ignore
        )

        super().__init__(**init_vals)
        if border_table.prop_has_attribs:
            self._set_style("border_table", border_table, *border_table.get_attrs())  # type: ignore
        if padding_fmt is not None:
            self._set_style("padding", padding_fmt, *padding_fmt.get_attrs())

        if shadow_fmt is not None:
            self._set_style("shadow", shadow_fmt)
        if diag_dn is not None:
            self._set_style("diag_dn", diag_dn)
        if diag_up is not None:
            self._set_style("diag_up", diag_up)

    # endregion init

    # region internal methods
    def _get_tb_cattribs(self) -> dict:
        return {
            "_property_name": self._props.tbl_border,
            "_supported_services_values": self._supported_services(),
            "_format_kind_prop": self.prop_format_kind,
            "_props_internal_attributes": StructBorderTableProps(
                left=self._props.tbl_bdr_left,
                top=self._props.tbl_bdr_top,
                right=self._props.tbl_bdr_right,
                bottom=self._props.tbl_bdr_bottom,
                horz=self._props.tbl_bdr_horz,
                vert=self._props.tbl_bdr_vert,
                dist=self._props.tbl_bdr_dist,
            ),
        }

    def _get_shadow_cattribs(self) -> dict:
        return {"_property_name": self._props.shadow, "_supported_services_values": self._supported_services()}

    def _get_padding_cattribs(self) -> dict:
        return {
            "_props_internal_attributes": BorderProps(
                left=self._props.pad_left,
                top=self._props.pad_top,
                right=self._props.pad_right,
                bottom=self._props.pad_btm,
            ),
            "_supported_services_values": self._supported_services(),
            "_format_kind_prop": self.prop_format_kind,
        }

    def _get_diagonal_up_cattribs(self) -> dict:
        return {
            "_property_name": self._props.diag_up,
            "_supported_services_values": self._supported_services(),
            "_format_kind_prop": self.prop_format_kind,
        }

    def _get_diagonal_dn_cattribs(self) -> dict:
        return {
            "_property_name": self._props.diag_dn,
            "_supported_services_values": self._supported_services(),
            "_format_kind_prop": self.prop_format_kind,
        }

    # endregion internal methods

    # region Overrides

    def _on_modifying(self, source: Any, event_args: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifying(source, event_args)

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.table.CellProperties",)
        return self._supported_services_values

    # region apply()
    @overload
    def apply(self, obj: Any) -> None:  # type: ignore
        ...

    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies padding to ``obj``

        Args:
            obj (object): Object that supports ``com.sun.star.style.ParagraphProperties`` service.

        Returns:
            None:
        """
        super().apply(obj, **kwargs)

    # endregion apply()

    # endregion Overrides

    # region Static Methods
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TBorders], obj: Any) -> _TBorders: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TBorders], obj: Any, **kwargs) -> _TBorders: ...

    @classmethod
    def from_obj(cls: Type[_TBorders], obj: Any, **kwargs) -> _TBorders:
        """
        Gets Borders instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedServiceError: If ``obj`` is not supported.

        Returns:
            Borders: Borders that represents ``obj`` borders.
        """
        # pylint: disable=protected-access
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        border_table = TableBorderStruct.from_obj(obj=obj, _cattribs=inst._get_tb_cattribs())
        if border_table.prop_has_attribs:
            inst._set_style("border_table", border_table, *border_table.get_attrs())
        else:
            inst._remove_style("border_table")

        shadow_fmt = Shadow.from_obj(obj=obj, _cattribs=inst._get_shadow_cattribs())
        inst._set_style("shadow", shadow_fmt)

        diag_dn = Side.from_obj(obj=obj, _cattribs=inst._get_diagonal_dn_cattribs())
        inst._set_style("diag_dn", diag_dn)

        diag_up = Side.from_obj(obj=obj, _cattribs=inst._get_diagonal_up_cattribs())
        inst._set_style("diag_up", diag_up)

        padding_fmt = Padding.from_obj(obj=obj, _cattribs=inst._get_padding_cattribs())
        inst._set_style("padding", padding_fmt, *padding_fmt.get_attrs())
        return inst

    # endregion from_obj()
    # endregion Static Methods

    # region Style Methods
    def _fmt_get_border_table(self: _TBorders, value: Side | None, side: str) -> Tuple[_TBorders, bool]:
        # pylint: disable=protected-access
        cp = self.copy()
        has_style = cp._has_style("border_table")

        if value is None:
            if has_style:
                cp._remove_style("border_table")
            return (cp, True)

        if not has_style:
            args = {side: value, "_cattribs": self._get_tb_cattribs()}
            border_table = TableBorderStruct(**args)
            cp._set_style("border_table", border_table)  # type: ignore
            return (cp, True)
        return (cp, False)

    def fmt_border_side(self: _TBorders, value: Side | None) -> _TBorders:
        """
        Gets copy of instance with left, right, top, bottom sides set or removed

        Args:
            value (~ooodev.format.inner.direct.structs.side.Side, optional): Side value

        Returns:
            Borders: Borders instance
        """
        # pylint: disable=protected-access
        cp, ret = self._fmt_get_border_table(value, "border_side")
        if ret:
            return cp
        bt = cast(TableBorderStruct, cp._get_style_inst("border_table"))
        bt.prop_left = value
        bt.prop_right = value
        bt.prop_top = value
        bt.prop_bottom = value
        cp._border_table = bt  # type: ignore
        return cp

    def fmt_left(self: _TBorders, value: Side | None) -> _TBorders:
        """
        Gets copy of instance with left set or removed

        Args:
            value (~ooodev.format.inner.direct.structs.side.Side, optional): Side value

        Returns:
            Borders: Borders instance
        """
        # pylint: disable=protected-access
        cp, ret = self._fmt_get_border_table(value, "left")
        if ret:
            return cp

        bt = cast(TableBorderStruct, cp._get_style_inst("border_table"))
        bt.prop_left = value
        return cp

    def fmt_right(self: _TBorders, value: Side | None) -> _TBorders:
        """
        Gets copy of instance with right set or removed

        Args:
            value (~ooodev.format.inner.direct.structs.side.Side, optional): Side value

        Returns:
            Borders: Borders instance
        """
        # pylint: disable=protected-access
        cp, ret = self._fmt_get_border_table(value, "right")
        if ret:
            return cp

        bt = cast(TableBorderStruct, cp._get_style_inst("border_table"))
        bt.prop_right = value
        return cp

    def fmt_top(self: _TBorders, value: Side | None) -> _TBorders:
        """
        Gets copy of instance with top set or removed

        Args:
            value (~ooodev.format.inner.direct.structs.side.Side, optional): Side value

        Returns:
            Borders: Borders instance
        """
        # pylint: disable=protected-access
        cp, ret = self._fmt_get_border_table(value, "top")
        if ret:
            return cp
        bt = cast(TableBorderStruct, cp._get_style_inst("border_table"))
        bt.prop_top = value
        return cp

    def fmt_bottom(self: _TBorders, value: Side | None) -> _TBorders:
        """
        Gets copy of instance with bottom set or removed

        Args:
            value (~ooodev.format.inner.direct.structs.side.Side, optional): Side value

        Returns:
            Borders: Borders instance
        """
        # pylint: disable=protected-access
        cp, ret = self._fmt_get_border_table(value, "bottom")
        if ret:
            return cp

        bt = cast(TableBorderStruct, cp._get_style_inst("border_table"))
        bt.prop_bottom = value
        return cp

    def fmt_horizontal(self: _TBorders, value: Side | None) -> _TBorders:
        """
        Gets copy of instance with horizontal set or removed

        Args:
            value (~ooodev.format.inner.direct.structs.side.Side, optional): Side value

        Returns:
            Borders: Borders instance
        """
        # pylint: disable=protected-access
        cp, ret = self._fmt_get_border_table(value, "horizontal")
        if ret:
            return cp

        bt = cast(TableBorderStruct, cp._get_style_inst("border_table"))
        bt.prop_horizontal = value
        return cp

    def fmt_vertical(self: _TBorders, value: Side | None) -> _TBorders:
        """
        Gets copy of instance with vertical set or removed

        Args:
            value (~ooodev.format.inner.direct.structs.side.Side, optional): Side value

        Returns:
            Borders: Borders instance
        """
        # pylint: disable=protected-access
        cp, ret = self._fmt_get_border_table(value, "vertical")
        if ret:
            return cp

        bt = cast(TableBorderStruct, cp._get_style_inst("border_table"))
        bt.prop_vertical = value
        return cp

    def fmt_distance(self: _TBorders, value: float | UnitT | None) -> _TBorders:
        """
        Gets copy of instance with distance set or removed

        Args:
            value (float, UnitT, optional): Distance value

        Returns:
            Borders: Borders instance
        """
        # pylint: disable=protected-access
        cp = self.copy()

        bt = cast(TableBorderStruct, cp._get_style_inst("distance"))
        bt.prop_distance = value
        return cp

    def fmt_diagonal_down(self: _TBorders, value: Side | None) -> _TBorders:
        """
        Gets copy of instance with diagonal down set or removed

        Args:
            value (~ooodev.format.inner.direct.structs.side.Side, optional): Side value

        Returns:
            Borders: Borders instance
        """
        # pylint: disable=protected-access
        cp = self.copy()
        if value is None:
            cp._remove_style("diag_dn")
        else:
            cp._set_style("diag_dn", value.copy(_cattribs=self._get_diagonal_dn_cattribs()))
        return cp

    def fmt_diagonal_up(self: _TBorders, value: Side | None) -> _TBorders:
        """
        Gets copy of instance with diagonal up set or removed

        Args:
            value (~ooodev.format.inner.direct.structs.side.Side, optional): Side value

        Returns:
            Borders: Borders instance
        """
        # pylint: disable=protected-access
        cp = self.copy()
        if value is None:
            cp._remove_style("diag_up")
        else:
            cp._set_style("diag_up", value.copy(_cattribs=self._get_diagonal_up_cattribs()))
        return cp

    def fmt_shadow(self: _TBorders, value: Shadow | None) -> _TBorders:
        """
        Gets copy of instance with shadow set or removed

        Args:
            value (Shadow, optional): Shadow value

        Returns:
            Borders: Borders instance
        """
        # pylint: disable=protected-access
        cp = self.copy()
        if value is None:
            cp._remove_style("shadow")
        else:
            shadow_fmt = value.copy(_cattribs=self._get_shadow_cattribs())
            cp._set_style("shadow", shadow_fmt)
        return cp

    def fmt_padding(self: _TBorders, value: Padding | None) -> _TBorders:
        """
        Gets copy of instance with padding set or removed

        Args:
            value (Padding, optional): Padding value

        Returns:
            Borders: Borders instance
        """
        # pylint: disable=protected-access
        cp = self.copy()
        if value is None:
            cp._remove_style("padding")
        else:
            cp._set_style("padding", value.copy(_cattribs=self._get_padding_cattribs()))
        return cp

    # endregion Style Methods

    # region Properties

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.CELL
        return self._format_kind_prop

    @property
    def prop_inner_padding(self) -> Padding:
        """Gets Padding instance"""
        try:
            return self._direct_inner_padding
        except AttributeError:
            self._direct_inner_padding = cast(Padding, self._get_style_inst("padding"))
        return self._direct_inner_padding

    @property
    def prop_inner_border_table(self) -> TableBorderStruct:
        """Gets border table instance"""
        try:
            return self._direct_inner_table
        except AttributeError:
            self._direct_inner_table = cast(TableBorderStruct, self._get_style_inst("border_table"))
        return self._direct_inner_table

    @property
    def prop_inner_shadow(self) -> Shadow | None:
        """Gets inner shadow instance"""
        try:
            return self._direct_inner_shadow
        except AttributeError:
            self._direct_inner_shadow = cast(Shadow, self._get_style_inst("shadow"))
        return self._direct_inner_shadow

    @property
    def prop_inner_diagonal_up(self) -> Side | None:
        """Gets inner Diagonal up instance"""
        try:
            return self._direct_dial_up
        except AttributeError:
            self._direct_dial_up = cast(Side, self._get_style_inst("diag_up"))
        return self._direct_dial_up

    @property
    def prop_inner_diagonal_dn(self) -> Side | None:
        """Gets inner Diagonal down instance"""
        try:
            return self._direct_dial_dn
        except AttributeError:
            self._direct_dial_dn = cast(Side, self._get_style_inst("diag_dn"))
        return self._direct_dial_dn

    @property
    def _props(self) -> CellBordersProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = CellBordersProps(
                tbl_border="TableBorder2",
                shadow="ShadowFormat",
                diag_up="DiagonalBLTR2",
                diag_dn="DiagonalTLBR2",
                pad_left="ParaLeftMargin",
                pad_top="ParaTopMargin",
                pad_right="ParaRightMargin",
                pad_btm="ParaBottomMargin",
                tbl_bdr_left=PropPair("LeftLine", "IsLeftLineValid"),
                tbl_bdr_top=PropPair("TopLine", "IsTopLineValid"),
                tbl_bdr_right=PropPair("RightLine", "IsRightLineValid"),
                tbl_bdr_bottom=PropPair("BottomLine", "IsBottomLineValid"),
                tbl_bdr_horz=PropPair("HorizontalLine", "IsHorizontalLineValid"),
                tbl_bdr_vert=PropPair("VerticalLine", "IsVerticalLineValid"),
                tbl_bdr_dist=PropPair("Distance", "IsDistanceValid"),
            )
        return self._props_internal_attributes

    @property
    def default(self: _TBorders) -> _TBorders:  # type: ignore[misc]
        """Gets Default Border."""
        # pylint: disable=protected-access
        # pylint: disable=unexpected-keyword-arg
        try:
            return self._default_inst
        except AttributeError:
            self._default_inst = self.__class__(
                border_side=Side(), padding=Padding(_cattribs=self._get_padding_cattribs()).default  # type: ignore
            )
            if self.has_update_obj():
                self._default_inst.set_update_obj(self.get_update_obj())
            self._default_inst._is_default_inst = True
        return self._default_inst

    @property
    def empty(self: _TBorders) -> _TBorders:  # type: ignore[misc]
        """Gets Empty Border. When style is applied formatting is removed."""
        # pylint: disable=protected-access
        # pylint: disable=unexpected-keyword-arg
        try:
            return self._empty_inst
        except AttributeError:
            side = Side()
            self._empty_inst = self.__class__(
                border_side=side.empty,
                vertical=side.empty,
                horizontal=side.empty,
                diagonal_down=side.empty,
                diagonal_up=side.empty,
                distance=0.0,
                shadow=Shadow(_cattribs=self._get_shadow_cattribs()).empty,  # type: ignore
                padding=Padding(_cattribs=self._get_padding_cattribs()).default,  # type: ignore
                _cattribs=self._get_internal_cattribs(),  # type: ignore
            )
            if self.has_update_obj():
                self._empty_inst.set_update_obj(self.get_update_obj())
            self._empty_inst._is_default_inst = True
        return self._empty_inst

    # endregion Properties
