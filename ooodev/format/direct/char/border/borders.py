"""
Module for managing character borders.

.. versionadded:: 0.9.0
"""
# region imports
from __future__ import annotations
from typing import Tuple, overload

import uno

from .....events.args.cancel_event_args import CancelEventArgs
from .....exceptions import ex as mEx
from .....meta.static_prop import static_prop
from .....utils import lo as mLo
from ....style_base import StyleMulti

from ....kind.border_kind import BorderKind
from ....kind.format_kind import FormatKind
from ...structs.side import Side as Side, SideFlags as SideFlags, LineSize as LineSize
from .shadow import Shadow as Shadow
from .padding import Padding as Padding
from .sides import Sides

from ooo.dyn.table.border_line import BorderLine as BorderLine
from ooo.dyn.table.border_line_style import BorderLineStyleEnum as BorderLineStyleEnum
from ooo.dyn.table.border_line2 import BorderLine2 as BorderLine2
from ooo.dyn.table.shadow_format import ShadowFormat as ShadowFormat
from ooo.dyn.table.shadow_location import ShadowLocation as ShadowLocation


# endregion imports


class Borders(StyleMulti):
    """
    Border used in styles for characters.

    All methods starting with ``fmt_`` can be used to chain together Borders properties.

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
        shadow: Shadow | None = None,
        padding: Padding | None = None,
    ) -> None:
        """
        Constructor

        Args:
            left (Side | None, optional): Determines the line style at the left edge.
            right (Side | None, optional): Determines the line style at the right edge.
            top (Side | None, optional): Determines the line style at the top edge.
            bottom (Side | None, optional): Determines the line style at the bottom edge.
            border_side (Side | None, optional): Determines the line style at the top, bottom, left, right edges. If this argument has a value then arguments ``top``, ``bottom``, ``left``, ``right`` are ignored
            shadow (BorderShadow | None, optional): Character Shadow
            padding (Padding | None, optional): Character padding
        """
        init_vals = {}

        sides = Sides(
            left=left,
            right=right,
            top=top,
            bottom=bottom,
            border_side=border_side,
        )

        super().__init__(**init_vals)

        if sides.prop_has_attribs:
            self._set_style("sides", sides, *sides.get_attrs())
        if not padding is None:
            self._set_style("padding", padding, *padding.get_attrs())
        if not shadow is None:
            self._set_style("shadow", shadow, *shadow.get_attrs())

    # endregion init
    def _set_side(self, side: Side | None, pos: BorderKind, inst: Borders | None = None) -> None:
        if inst is None:
            inst = self
        if pos == BorderKind.NONE:
            inst._remove_style("sides")
            return
        sides_info = inst._get_style("sides")
        if sides_info is None:
            sides = Sides()
        else:
            sides = sides_info[0].copy()
        if BorderKind.BOTTOM in pos:
            sides.prop_bottom = side
        if BorderKind.LEFT in pos:
            sides.prop_left = side
        if BorderKind.RIGHT in pos:
            sides.prop_right = side
        if BorderKind.TOP in pos:
            sides.prop_bottom = side
        inst._set_style("sides", sides, *sides.get_attrs())

    # region format Methods
    def fmt_border_side(self, value: Side | None) -> Borders:
        """
        Gets copy of instance with left, right, top, bottom sides set or removed

        Args:
            value (Side | None): Side value

        Returns:
            Borders: Borders instance
        """
        cp = self.copy()
        if value is None:
            self._set_side(side=value, pos=BorderKind.NONE, inst=cp)
            return cp
        self._set_side(side=value, pos=BorderKind.ALL, inst=cp)
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
        if value is None:
            self._set_side(side=value, pos=BorderKind.NONE, inst=cp)
            return cp
        self._set_side(side=value, pos=BorderKind.LEFT, inst=cp)
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
        if value is None:
            self._set_side(side=value, pos=BorderKind.NONE, inst=cp)
            return cp
        self._set_side(side=value, pos=BorderKind.RIGHT, inst=cp)
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
        if value is None:
            self._set_side(side=value, pos=BorderKind.NONE, inst=cp)
            return cp
        self._set_side(side=value, pos=BorderKind.TOP, inst=cp)
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
        if value is None:
            self._set_side(side=value, pos=BorderKind.NONE, inst=cp)
            return cp
        self._set_side(side=value, pos=BorderKind.BOTTOM, inst=cp)
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
            cp._remove_style("shadow")
        else:
            cp._set_style("shadow", value, *value.get_attrs())
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
        if value is None:
            cp._remove_style("padding")
        else:
            cp._set_style("padding", value, *value.get_attrs())
        return cp

    # endregion format Methods

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        """
        Gets a tuple of supported services (``com.sun.star.style.CharacterProperties``,)

        Returns:
            Tuple[str, ...]: Supported services
        """
        return ("com.sun.star.style.CharacterProperties",)

    def _on_modifing(self, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifing(event)

    # region apply()
    @overload
    def apply(self, obj: object) -> None:
        ...

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies padding to ``obj``

        Args:
            obj (object): Object that supports ``com.sun.star.style.CharacterProperties`` service.

        Returns:
            None:
        """

        try:
            super().apply(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion apply()
    # endregion methods

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.CHAR

    @static_prop
    def default() -> Borders:  # type: ignore[misc]
        """Gets Default Border. Static Property"""
        try:
            return Borders._DEFAULT_INST
        except AttributeError:
            Borders._DEFAULT_INST = Borders(border_side=Side.empty, padding=Padding.default, shadow=Shadow.empty)
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
                shadow=Shadow.empty,
                padding=Padding.default,
            )
            Borders._EMPTY_INST._is_default_inst = True
        return Borders._EMPTY_INST

    # endregion Properties
