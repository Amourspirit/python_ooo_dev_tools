"""
Module for managing table borders (cells and ranges).

.. versionadded:: 0.9.0
"""
# region imports
from __future__ import annotations
from typing import Any, overload, cast, Tuple, TypeVar

import uno

from .....events.args.cancel_event_args import CancelEventArgs
from .....exceptions import ex as mEx
from .....meta.static_prop import static_prop
from .....utils import lo as mLo
from ....style_base import StyleMulti

from ....kind.format_kind import FormatKind
from .padding import Padding as Padding
from ...structs.table_border_struct import TableBorderStruct
from .shadow import Shadow
from ...structs.side import Side as Side, BorderLineKind as BorderLineKind

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
        distance: float | None = None,
        diagonal_down: Side | None = None,
        diagonal_up: Side | None = None,
        shadow: Shadow | None = None,
        padding: Padding | None = None,
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
            init_vals["ShadowFormat"] = shadow.get_uno_struct()
        if not diagonal_down is None:
            init_vals["DiagonalTLBR2"] = diagonal_down.get_uno_struct()
        if not diagonal_up is None:
            init_vals["DiagonalBLTR2"] = diagonal_up.get_uno_struct()

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

        super().__init__(**init_vals)
        if border_table.prop_has_attribs:
            self._set_style("border_table", border_table, *border_table.get_attrs())
        if not padding is None:
            self._set_style("padding", padding, *padding.get_attrs())

    # endregion init

    # region methods

    def _get_tb_cattribs(self) -> dict:
        return {"_property_name": "TableBorder2", "_supported_services_values": self._supported_services()}

    def _on_modifing(self, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifing(event)

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.table.CellProperties",)
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

    # endregion methods

    # region Style Methods
    def fmt_border_side(self: _TBorders, value: Side | None) -> _TBorders:
        """
        Gets copy of instance with left, right, top, bottom sides set or removed

        Args:
            value (Side | None): Side value

        Returns:
            Borders: Borders instance
        """
        cp = self.copy()
        if cp._border_table is None and value is None:
            return cp
        if cp._border_table is None:
            cp._border_table = TableBorderStruct(border_side=value, _cattribs=self._get_tb_cattribs())
            return cp
        bt = cp._border_table.copy()
        bt.prop_left = value
        bt.prop_right = value
        bt.prop_top = value
        bt.prop_bottom = value
        cp._border_table = bt
        return cp

    def fmt_left(self: _TBorders, value: Side | None) -> _TBorders:
        """
        Gets copy of instance with left set or removed

        Args:
            value (Side | None): Side value

        Returns:
            Borders: Borders instance
        """
        cp = self.copy()
        if cp._border_table is None and value is None:
            return cp
        if cp._border_table is None:
            cp._border_table = TableBorderStruct(left=value, _cattribs=self._get_tb_cattribs())
            return cp
        bt = cp._border_table.copy()
        bt.prop_left = value
        cp._border_table = bt
        return cp

    def fmt_right(self: _TBorders, value: Side | None) -> _TBorders:
        """
        Gets copy of instance with right set or removed

        Args:
            value (Side | None): Side value

        Returns:
            Borders: Borders instance
        """
        cp = self.copy()
        if cp._border_table is None and value is None:
            return cp
        if cp._border_table is None:
            cp._border_table = TableBorderStruct(right=value, _cattribs=self._get_tb_cattribs())
            return cp
        bt = cp._border_table.copy()
        bt.prop_right = value
        cp._border_table = bt
        return cp

    def fmt_top(self: _TBorders, value: Side | None) -> _TBorders:
        """
        Gets copy of instance with top set or removed

        Args:
            value (Side | None): Side value

        Returns:
            Borders: Borders instance
        """
        cp = self.copy()
        if cp._border_table is None and value is None:
            return cp
        if cp._border_table is None:
            cp._border_table = TableBorderStruct(top=value, _cattribs=self._get_tb_cattribs())
            return cp
        bt = cp._border_table.copy()
        bt.prop_top = value
        cp._border_table = bt
        return cp

    def fmt_bottom(self: _TBorders, value: Side | None) -> _TBorders:
        """
        Gets copy of instance with bottom set or removed

        Args:
            value (Side | None): Side value

        Returns:
            Borders: Borders instance
        """
        cp = self.copy()
        if cp._border_table is None and value is None:
            return cp
        if cp._border_table is None:
            cp._border_table = TableBorderStruct(bottom=value, _cattribs=self._get_tb_cattribs())
            return cp
        bt = cp._border_table.copy()
        bt.prop_bottom = value
        cp._border_table = bt
        return cp

    def fmt_horizontal(self: _TBorders, value: Side | None) -> _TBorders:
        """
        Gets copy of instance with horizontal set or removed

        Args:
            value (Side | None): Side value

        Returns:
            Borders: Borders instance
        """
        cp = self.copy()
        if cp._border_table is None and value is None:
            return cp
        if cp._border_table is None:
            cp._border_table = TableBorderStruct(horizontal=value, _cattribs=self._get_tb_cattribs())
            return cp
        bt = cp._border_table.copy()
        bt.prop_horizontal = value
        cp._border_table = bt
        return cp

    def fmt_vertical(self: _TBorders, value: Side | None) -> _TBorders:
        """
        Gets copy of instance with vertical set or removed

        Args:
            value (Side | None): Side value

        Returns:
            Borders: Borders instance
        """
        cp = self.copy()
        if cp._border_table is None and value is None:
            return cp
        if cp._border_table is None:
            cp._border_table = TableBorderStruct(vertical=value, _cattribs=self._get_tb_cattribs())
            return cp
        bt = cp._border_table.copy()
        bt.prop_vertical = value
        cp._border_table = bt
        return cp

    def fmt_distance(self: _TBorders, value: float | None) -> _TBorders:
        """
        Gets copy of instance with distance set or removed

        Args:
            value (float | None): Distance value

        Returns:
            Borders: Borders instance
        """
        cp = self.copy()
        if cp._border_table is None and value is None:
            return cp
        if cp._border_table is None:
            cp._border_table = TableBorderStruct(distance=value, _cattribs=self._get_tb_cattribs())
            return cp
        bt = cp._border_table.copy()
        bt.prop_distance = value
        cp._border_table = bt
        return cp

    def fmt_diagonal_down(self: _TBorders, value: Side | None) -> _TBorders:
        """
        Gets copy of instance with diagonal down set or removed

        Args:
            value (Shadow | None): Side value

        Returns:
            Borders: Borders instance
        """
        cp = self.copy()
        if value is None:
            cp._remove("DiagonalTLBR2")
        else:
            cp._set("DiagonalTLBR2", value.get_uno_struct())
        return cp

    def fmt_diagonal_up(self: _TBorders, value: Side | None) -> _TBorders:
        """
        Gets copy of instance with diagonal up set or removed

        Args:
            value (Shadow | None): Side value

        Returns:
            Borders: Borders instance
        """
        cp = self.copy()
        if value is None:
            cp._remove("DiagonalBLTR2")
        else:
            cp._set("DiagonalBLTR2", value.get_uno_struct())
        return cp

    def fmt_shadow(self: _TBorders, value: Shadow | None) -> _TBorders:
        """
        Gets copy of instance with shadow set or removed

        Args:
            value (Shadow | None): Shadow value

        Returns:
            Borders: Borders instance
        """
        cp = self.copy()
        if value is None:
            cp._remove("ShadowFormat")
        else:
            cp._set("ShadowFormat", value.get_uno_struct())
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
        cp._padding = value
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
