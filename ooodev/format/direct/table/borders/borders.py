"""
Module for managing table borders (cells and ranges).

.. versionadded:: 0.9.0
"""
# region imports
from __future__ import annotations
from typing import Any, Type, overload, cast, Tuple, TypeVar

import uno

from .....events.args.cancel_event_args import CancelEventArgs
from .....exceptions import ex as mEx
from .....meta.static_prop import static_prop
from .....proto.unit_obj import UnitObj
from .....utils import lo as mLo
from ....kind.format_kind import FormatKind
from ....style_base import StyleMulti
from ...cell.border.padding import Padding
from ...cell.border.shadow import Shadow
from ...structs.side import Side as Side, BorderLineKind as BorderLineKind
from ...structs.table_border_struct import TableBorderStruct
from ...structs.table_border_distances_struct import TableBorderDistancesStruct
from ...common.props.table_borders_props import TableBordersProps
from .....utils import props as mProps

from ooo.dyn.table.border_line import BorderLine as BorderLine
from ooo.dyn.table.border_line2 import BorderLine2 as BorderLine2
from ooo.dyn.table.shadow_format import ShadowFormat as ShadowFormat
from ooo.dyn.table.shadow_location import ShadowLocation as ShadowLocation

# endregion imports

_TBorders = TypeVar(name="_TBorders", bound="Borders")


class Borders(StyleMulti):
    """
    Table Borders used in styles for table cells and ranges.

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

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
        distance: float | UnitObj | None = None,
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
            distance (float, UnitObj, optional): Contains the distance between the lines and other contents in ``mm`` units or :ref:`proto_unit_obj`.
            shadow (Shadow, optional): Cell Shadow.
            padding (BorderPadding, optional): Cell padding.
            merge_adjacent (bool, optional): Specifies if adjacent line style are to be merged.
        """
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
            _cattribs=self._get_tb_cattribs(),
        )

        if not padding is None:
            # using paddint and converting to TableBorderDistancesStruct.
            # TableBorderDistancesStruct could be used but padding is used on other places so keep it that same.
            tb_padding_struct = TableBorderDistancesStruct(
                left=padding.prop_left or 0,
                right=padding.prop_right or 0,
                top=padding.prop_top or 0,
                bottom=padding.prop_bottom or 0,
                _cattribs=self._get_tbd_cattribs(),
            )
        else:
            tb_padding_struct = None

        if not merge_adjacent is None:
            self.prop_merge_adjacent = merge_adjacent
        if border_table.prop_has_attribs:
            self._set_style("border_table", border_table)
        if not tb_padding_struct is None:
            self._set_style("padding", tb_padding_struct)
        if not shadow_fmt is None:
            self._set_style("shadow", shadow_fmt)

    # endregion init

    # region Internal Methods
    def _get_tb_cattribs(self) -> dict:
        return {"_property_name": self._props.tbl_border, "_supported_services_values": self._supported_services()}

    def _get_shadow_cattribs(self) -> dict:
        return {"_property_name": self._props.shadow, "_supported_services_values": self._supported_services()}

    def _get_tbd_cattribs(self) -> dict:
        return {"_property_name": self._props.tbl_distance, "_supported_services_values": self._supported_services()}

    # endregion Internal Methods

    # region Overrides

    def _on_modifing(self, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifing(event)

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.text.TextTable",)
        return self._supported_services_values

    # region apply()
    @overload
    def apply(self, obj: object) -> None:
        ...

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies padding to ``obj``

        Args:
            obj (object): Object that supports ``com.sun.star.style.ParagraphProperties`` service.

        Returns:
            None:
        """
        super().apply(obj, **kwargs)

    # endregion apply()

    def _props_set(self, obj: object, **kwargs: Any) -> None:
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
    def from_obj(cls: Type[_TBorders], obj: object) -> _TBorders:
        ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TBorders], obj: object, **kwargs) -> _TBorders:
        ...

    @classmethod
    def from_obj(cls: Type[_TBorders], obj: object, **kwargs) -> _TBorders:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Borders: ``Borders`` instance that represents ``obj`` border properties.
        """
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')
        inst.prop_merge_adjacent = bool(mProps.Props.get(obj, inst._props.merge, True))
        tbl_border = TableBorderStruct.from_obj(obj, _cattribs=inst._get_tb_cattribs())
        tbl_border_disance = TableBorderDistancesStruct.from_obj(obj, _cattribs=inst._get_tbd_cattribs())
        shadow_fmt = Shadow.from_obj(obj, _cattribs=inst._get_shadow_cattribs())
        inst._set_style("border_table", tbl_border)
        inst._set_style("padding", tbl_border_disance)
        inst._set_style("shadow", shadow_fmt)
        return inst

    # endregion from_obj()

    # endregion Static Methods

    # region Style Methods
    def _fmt_get_border_table(self: _TBorders, value: Side | None, side: str) -> Tuple[_TBorders, bool]:
        cp = self.copy()
        has_style = cp._has_style("border_table")

        if value is None:
            if has_style:
                cp._remove_style("border_table")
            return (cp, True)

        if not has_style:
            args = {side: value, "_cattribs": self._get_tb_cattribs()}
            border_table = TableBorderStruct(**args)
            cp._set_style("border_table", border_table)
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
        cp, ret = self._fmt_get_border_table(value, "vertical")
        if ret:
            return cp

        bt = cast(TableBorderStruct, cp._get_style_inst("border_table"))
        bt.prop_vertical = value
        return cp

    def fmt_distance(self: _TBorders, value: float | UnitObj | None) -> _TBorders:
        """
        Gets copy of instance with distance set or removed

        Args:
            value (float, UnitObj, optional): Distance value

        Returns:
            Borders: Borders instance
        """
        cp, ret = self._fmt_get_border_table(value, "vertical")
        if ret:
            return cp

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
            value (Padding | None): Padding value

        Returns:
            Borders: Borders instance
        """
        cp = self.copy()
        if value is None:
            cp._remove_style("padding")
        else:
            tb_padding_struct = TableBorderDistancesStruct(
                left=value.prop_left or 0,
                right=value.prop_right or 0,
                top=value.prop_top or 0,
                bottom=value.prop_bottom or 0,
                _cattribs=self._get_tbd_cattribs(),
            )
            cp._set_style("padding", tb_padding_struct)
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
        """Gets innner Padding as ``TableBorderDistancesStruct`` instance"""
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
            )
        return self._props_internal_attributes

    @static_prop
    def default() -> Borders:  # type: ignore[misc]
        """Gets Default Border. Static Property"""
        try:
            return Borders._DEFAULT_INST
        except AttributeError:
            Borders._DEFAULT_INST = Borders(border_side=Side(), padding=Padding.default)
            Borders._DEFAULT_INST._is_default_inst = True
        return Borders._DEFAULT_INST

    @static_prop
    def empty() -> Borders:  # type: ignore[misc]
        """Gets Empty Border. Static Property. When style is applied formatting is removed."""
        try:
            return Borders._EMPTY_INST
        except AttributeError:
            Borders._EMPTY_INST = Borders(
                border_side=Side.empty,
                vertical=Side.empty,
                horizontal=Side.empty,
                diagonal_down=Side.empty,
                diagonal_up=Side.empty,
                distance=0.0,
                shadow=Shadow.empty,
                padding=Padding.default,
            )
            Borders._EMPTY_INST._is_default_inst = True
        return Borders._EMPTY_INST

    # endregion Properties
