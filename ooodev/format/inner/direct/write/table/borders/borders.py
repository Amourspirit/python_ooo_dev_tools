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
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleMulti
from ooodev.format.inner.direct.calc.border.padding import Padding
from ooodev.format.inner.direct.calc.border.shadow import Shadow
from ooodev.format.inner.common.props.prop_pair import PropPair
from ooodev.format.inner.common.props.struct_border_table_props import StructBorderTableProps
from ooodev.format.inner.common.props.table_borders_props import TableBordersProps
from ooodev.format.inner.direct.structs.side import Side
from ooodev.format.inner.direct.structs.table_border_distances_struct import TableBorderDistancesStruct
from ooodev.format.inner.direct.structs.table_border_struct import TableBorderStruct


# endregion imports

_TBorders = TypeVar(name="_TBorders", bound="Borders")


class Borders(StyleMulti):
    """
    Table Borders used in styles for table cells and ranges.

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. seealso::

        - :ref:`help_writer_format_direct_table_borders`

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
        shadow: Shadow | None = None,
        padding: Padding | None = None,
        merge_adjacent: bool | None = None,
    ) -> None:
        """
        Constructor

        Args:
            left (Side,, optional): Specifies the line style at the left edge.
            right (Side, optional): Specifies the line style at the right edge.
            top (Side, optional): Specifies the line style at the top edge.
            bottom (Side, optional): Specifies the line style at the bottom edge.
            border_side (Side, optional): Specifies the line style at the top, bottom, left, right edges. If this argument has a value then arguments ``top``, ``bottom``, ``left``, ``right`` are ignored
            horizontal (Side, optional): Specifies the line style of horizontal lines for the inner part of a cell range.
            vertical (Side, optional): Specifies the line style of vertical lines for the inner part of a cell range.
            distance (float, UnitT, optional): Contains the distance between the lines and other contents in ``mm`` units or :ref:`proto_unit_obj`.
            shadow (Shadow, optional): Cell Shadow.
            padding (BorderPadding, optional): Cell padding.
            merge_adjacent (bool, optional): Specifies if adjacent line style are to be merged.

        Returns:
            None:

        Hint:
            - ``LineSize`` can be imported from ``ooodev.format.writer.direct.char.borders``
            - ``Padding`` can be imported from ``ooodev.format.inner.direct.calc.border.padding``
            - ``ShadowFormat`` can be imported from ``ooodev.format.writer.direct.char.borders``
            - ``Side`` can be imported from ``ooodev.format.writer.direct.char.borders``
            - ``side`` can be imported from ``ooodev.format.inner.direct.structs.side``
            - ``ShadowLocation`` can be imported ``from ooo.dyn.table.shadow_location``
            - ``Shadow`` can be imported from ``ooodev.format.inner.direct.calc.border.shadow``

        See Also:
            - :ref:`help_writer_format_direct_table_borders`
        """
        # pylint: disable=unexpected-keyword-arg
        super().__init__()

        if shadow is None:
            shadow_fmt = None
        else:
            shadow_fmt = shadow.copy(_cattribs=self._get_shadow_cattribs())

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

        if padding is not None:
            # using paddint and converting to TableBorderDistancesStruct.
            # TableBorderDistancesStruct could be used but padding is used on other places so keep it that same.
            tb_padding_struct = TableBorderDistancesStruct(
                left=padding.prop_left or 0,
                right=padding.prop_right or 0,
                top=padding.prop_top or 0,
                bottom=padding.prop_bottom or 0,
                _cattribs=self._get_tbd_cattribs(),  # type: ignore
            )
        else:
            tb_padding_struct = None

        if merge_adjacent is not None:
            self.prop_merge_adjacent = merge_adjacent
        if border_table.prop_has_attribs:
            self._set_style("border_table", border_table)  # type: ignore
        if tb_padding_struct is not None:
            self._set_style("padding", tb_padding_struct)  # type: ignore
        if shadow_fmt is not None:
            self._set_style("shadow", shadow_fmt)

    # endregion init

    # region Internal Methods
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
        return {
            "_property_name": self._props.shadow,
            "_supported_services_values": self._supported_services(),
            "_format_kind_prop": self.prop_format_kind,
        }

    def _get_tbd_cattribs(self) -> dict:
        return {
            "_property_name": self._props.tbl_distance,
            "_supported_services_values": self._supported_services(),
            "_format_kind_prop": self.prop_format_kind,
        }

    # endregion Internal Methods

    # region Overrides

    def _on_modifying(self, source: Any, event_args: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifying(source, event_args)

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.text.TextTable",)
        return self._supported_services_values

    # region apply()
    @overload
    def apply(self, obj: Any) -> None: ...

    @overload
    def apply(self, obj: Any, **kwargs) -> None: ...

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

    def _props_set(self, obj: Any, **kwargs: Any) -> None:
        try:
            super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

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
        Gets instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Borders: ``Borders`` instance that represents ``obj`` border properties.
        """
        # pylint: disable=protected-access
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')
        inst.prop_merge_adjacent = bool(mProps.Props.get(obj, inst._props.merge, True))
        tbl_border = TableBorderStruct.from_obj(obj, _cattribs=inst._get_tb_cattribs())
        tbl_border_distance = TableBorderDistancesStruct.from_obj(obj, _cattribs=inst._get_tbd_cattribs())
        shadow_fmt = Shadow.from_obj(obj, _cattribs=inst._get_shadow_cattribs())
        inst._set_style("border_table", tbl_border)
        inst._set_style("padding", tbl_border_distance)
        inst._set_style("shadow", shadow_fmt)
        inst.set_update_obj(obj)
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
            value (Side, optional): Side value

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
        return cp

    def fmt_left(self: _TBorders, value: Side | None) -> _TBorders:
        """
        Gets copy of instance with left set or removed

        Args:
            value (Side, optional): Side value

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
            value (Side, optional): Side value

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
            value (Side, optional): Side value

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
            value (Side, optional): Side value

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
            value (Side, optional): Side value

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
            value (Side, optional): Side value

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
        # pylint: disable=unexpected-keyword-arg
        cp = self.copy()
        if value is None:
            cp._remove_style("padding")
        else:
            tb_padding_struct = TableBorderDistancesStruct(
                left=value.prop_left or 0,
                right=value.prop_right or 0,
                top=value.prop_top or 0,
                bottom=value.prop_bottom or 0,
                _cattribs=self._get_tbd_cattribs(),  # type: ignore
            )
            cp._set_style("padding", tb_padding_struct)  # type: ignore
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
    def prop_merge_adjacent(self) -> bool | None:
        """
        Gets/Sets merge_adjacent.
        """
        return self._get(self._props.merge)

    @prop_merge_adjacent.setter
    def prop_merge_adjacent(self, value: bool | None):
        if value is None:
            self._remove(self._props.merge)
            return
        self._set(self._props.merge, value)

    @property
    def prop_inner_padding(self) -> TableBorderDistancesStruct:
        """Gets inner Padding as ``TableBorderDistancesStruct`` instance"""
        try:
            return self._direct_inner_padding
        except AttributeError:
            self._direct_inner_padding = cast(TableBorderDistancesStruct, self._get_style_inst("padding"))
        return self._direct_inner_padding

    @property
    def prop_inner_border_table(self) -> TableBorderStruct:
        """Gets inner border table instance"""
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
    def _props(self) -> TableBordersProps:
        """Gets _props."""
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = TableBordersProps(
                tbl_border="TableBorder2",
                shadow="ShadowFormat",
                tbl_distance="TableBorderDistances",
                merge="CollapsingBorders",
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
                border_side=Side(), padding=Padding(_cattribs=self._get_tbd_cattribs()).default  # type: ignore
            )
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
                diagonal_down=side.empty,  # type: ignore
                diagonal_up=side.empty,  # type: ignore
                distance=0.0,
                shadow=Shadow(_cattribs=self._get_shadow_cattribs()).empty,  # type: ignore
                padding=Padding(_cattribs=self._get_tbd_cattribs()).default,  # type: ignore
            )
            self._empty_inst._is_default_inst = True
        return self._empty_inst

    # endregion Properties
