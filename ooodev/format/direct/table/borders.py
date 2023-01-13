"""
Module for managing table borders (cells and ranges).

.. versionadded:: 0.9.0
"""
# region imports
from __future__ import annotations
from ast import Tuple
from typing import overload

import uno
from ....exceptions import ex as mEx
from ....meta.static_prop import static_prop
from ....utils import lo as mLo
from ...style_base import StyleMulti

from ..structs import side
from ..structs import shadow
from ..structs import border_table
from ..para import padding
from ..structs.side import Side as Side, SideFlags as SideFlags
from ..structs.shadow import Shadow
from ..structs.border_table import BorderTable as BorderTable
from ..para.padding import Padding as Padding
from ...kind.style_kind import StyleKind

from ooo.dyn.table.border_line import BorderLine as BorderLine
from ooo.dyn.table.border_line_style import BorderLineStyleEnum as BorderLineStyleEnum
from ooo.dyn.table.border_line2 import BorderLine2 as BorderLine2
from ooo.dyn.table.shadow_format import ShadowFormat as ShadowFormat
from ooo.dyn.table.shadow_location import ShadowLocation as ShadowLocation


# endregion imports


class Borders(StyleMulti):
    """
    Table Borders used in styles for table cells and ranges.

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

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

        super().__init__(**init_vals)
        if border_table.prop_has_attribs:
            self._set_style("border_table", border_table, *border_table.get_attrs())
        if not padding is None:
            self._set_style("padding", padding, *padding.get_attrs())

    # endregion init

    # region methods

    def _supported_services(self) -> Tuple[str, ...]:
        """
        Gets a tuple of supported services (``com.sun.star.table.CellProperties``,)

        Returns:
            Tuple[str, ...]: Supported services
        """
        return ("com.sun.star.table.CellProperties",)

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
        try:
            super().apply(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"CellBorder.apply_style(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion apply()

    # endregion methods

    # region Style Methods
    def fmt_border_side(self, value: Side | None) -> Borders:
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
            cp._border_table = BorderTable(border_side=value)
            return cp
        bt = cp._border_table.copy()
        bt.prop_left = value
        bt.prop_right = value
        bt.prop_top = value
        bt.prop_bottom = value
        cp._border_table = bt
        return cp

    def fmt_left(self, value: Side | None) -> Borders:
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
            cp._border_table = BorderTable(left=value)
            return cp
        bt = cp._border_table.copy()
        bt.prop_left = value
        cp._border_table = bt
        return cp

    def fmt_right(self, value: Side | None) -> Borders:
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
            cp._border_table = BorderTable(right=value)
            return cp
        bt = cp._border_table.copy()
        bt.prop_right = value
        cp._border_table = bt
        return cp

    def fmt_top(self, value: Side | None) -> Borders:
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
            cp._border_table = BorderTable(top=value)
            return cp
        bt = cp._border_table.copy()
        bt.prop_top = value
        cp._border_table = bt
        return cp

    def fmt_bottom(self, value: Side | None) -> Borders:
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
            cp._border_table = BorderTable(bottom=value)
            return cp
        bt = cp._border_table.copy()
        bt.prop_bottom = value
        cp._border_table = bt
        return cp

    def fmt_horizontal(self, value: Side | None) -> Borders:
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
            cp._border_table = BorderTable(horizontal=value)
            return cp
        bt = cp._border_table.copy()
        bt.prop_horizontal = value
        cp._border_table = bt
        return cp

    def fmt_vertical(self, value: Side | None) -> Borders:
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
            cp._border_table = BorderTable(vertical=value)
            return cp
        bt = cp._border_table.copy()
        bt.prop_vertical = value
        cp._border_table = bt
        return cp

    def fmt_distance(self, value: float | None) -> Borders:
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
            cp._border_table = BorderTable(distance=value)
            return cp
        bt = cp._border_table.copy()
        bt.prop_distance = value
        cp._border_table = bt
        return cp

    def fmt_diagonal_down(self, value: Side | None) -> Borders:
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
            cp._set("DiagonalTLBR2", value.get_border_line2())
        return cp

    def fmt_diagonal_up(self, value: Side | None) -> Borders:
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
            cp._set("DiagonalBLTR2", value.get_border_line2())
        return cp

    def fmt_shadow(self, value: Shadow | None) -> Borders:
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
            cp._set("ShadowFormat", value.get_shadow_format())
        return cp

    def fmt_padding(self, value: Padding | None) -> Borders:
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
    def prop_style_kind(self) -> StyleKind:
        """Gets the kind of style"""
        return StyleKind.CELL

    @static_prop
    def default() -> Borders:  # type: ignore[misc]
        """Gets Default Border. Static Property"""
        if Borders._DEFAULT is None:
            Borders._DEFAULT = Borders(border_side=Side(), padding=Padding.default)
        return Borders._DEFAULT

    @static_prop
    def empty() -> Borders:  # type: ignore[misc]
        """Gets Empty Border. Static Property. When style is applied formatting is removed."""
        if Borders._EMPTY is None:
            Borders._EMPTY = Borders(
                border_side=Side.empty,
                vertical=Side.empty,
                horizontal=Side.empty,
                diagonal_down=Side.empty,
                diagonal_up=Side.empty,
                distance=0.0,
                shadow=Shadow.empty,
                padding=Padding.default,
            )
        return Borders._EMPTY

    # endregion Properties
