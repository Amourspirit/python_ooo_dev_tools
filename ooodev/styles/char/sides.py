"""
Module for managing character border side.

.. versionadded:: 0.9.0
"""
# region imports
from __future__ import annotations
from typing import Tuple, cast, overload, TYPE_CHECKING

import uno
from ...events.args.key_val_cancel_args import KeyValCancelArgs
from ...exceptions import ex as mEx
from ...utils import info as mInfo
from ...utils import lo as mLo
from ...utils import props as mProps
from ..kind.style_kind import StyleKind
from ..structs import side
from ..structs.side import Side as Side, BorderLineStyleEnum as BorderLineStyleEnum
from ..style_base import StyleBase

from ooo.dyn.table.border_line2 import BorderLine2

if TYPE_CHECKING:
    from com.sun.star.style import CharacterProperties  # service

# endregion imports


class Sides(StyleBase):
    """
    Character Border for use in styles.

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``style_`` can be used to chain together Sides properties.

    .. versionadded:: 0.9.0
    """

    _CHAR_BORDERS = ("CharTopBorder", "CharBottomBorder", "CharLeftBorder", "CharRightBorder")
    # region init

    def __init__(
        self,
        left: Side | None = None,
        right: Side | None = None,
        top: Side | None = None,
        bottom: Side | None = None,
        border_side: Side | None = None,
    ) -> None:
        """
        Constructor

        Args:
            left (Side | None, optional): Determines the line style at the left edge.
            right (Side | None, optional): Determines the line style at the right edge.
            top (Side | None, optional): Determines the line style at the top edge.
            bottom (Side | None, optional): Determines the line style at the bottom edge.
            border_side (Side | None, optional): Determines the line style at the top, bottom, left, right edges. If this argument has a value then arguments ``top``, ``bottom``, ``left``, ``right`` are ignored
        """
        init_vals = {}
        if not border_side is None:
            init_vals["CharTopBorder"] = border_side
            init_vals["CharBottomBorder"] = border_side
            init_vals["CharLeftBorder"] = border_side
            init_vals["CharRightBorder"] = border_side

        else:
            if not top is None:
                init_vals["CharTopBorder"] = top
            if not bottom is None:
                init_vals["CharBottomBorder"] = bottom
            if not left is None:
                init_vals["CharLeftBorder"] = left
            if not right is None:
                init_vals["CharRightBorder"] = right

        self._has_attribs = len(init_vals) > 0
        super().__init__(**init_vals)

    # endregion init

    # region style methods
    # def style_left(self, value: Side | None = None) -> Sides:
    #     pass
    # endregion style methods
    # region methods

    def _supported_services(self) -> Tuple[str, ...]:
        """
        Gets a tuple of supported services (``com.sun.star.style.CharacterProperties``,)

        Returns:
            Tuple[str, ...]: Supported services
        """
        return ("com.sun.star.style.CharacterProperties",)

    # region apply_style()

    @overload
    def apply_style(self, obj: object) -> None:
        ...

    def apply_style(self, obj: object, **kwargs) -> None:
        """
        Applies Style to obj

        Args:
            obj (object): UNO object

        Returns:
            None:
        """
        try:
            super().apply_style(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"BorderChar.apply_style(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion apply_style()

    def on_property_setting(self, event_args: KeyValCancelArgs):
        """
        Raise for each property that is set

        Args:
            event_args (KeyValueCancelArgs): Event Args
        """
        if event_args.has("sides_border2_set"):
            return
        side = cast(Side, event_args.value)
        event_args.value = side.get_border_line2()
        event_args.set("sides_border2_set", True)

    @staticmethod
    def from_obj(obj: object) -> Sides:
        """
        Gets instance from object properties

        Args:
            obj (object): UNO object that has supports ``com.sun.star.style.CharacterProperties`` service.

        Raises:
            NotSupportedServiceError: If ``obj`` does not support ``com.sun.star.style.CharacterProperties`` service.
            PropertyNotFoundError: If ``obj`` does not have ``TableBorder2`` property.

        Returns:
            BorderTable: Border Table.
        """
        if not mInfo.Info.support_service(obj, "com.sun.star.style.CharacterProperties"):
            mLo.Lo.print('BorderChar.apply_style(): "com.sun.star.style.CharacterProperties" not supported')
            return

        bc = Sides()
        if bc._is_valid_service(obj):
            cp = cast("CharacterProperties", obj)
            empty = BorderLine2()
            for attr in Sides._CHAR_BORDERS:
                b2 = cast(BorderLine2, getattr(cp, attr, empty))
                side = Side.from_border2(b2)
                bc._set(attr, side)
        else:
            raise mEx.NotSupportedServiceError(bc._supported_services()[0])
        return bc

    # endregion methods

    # region style methods
    def style_border_side(self, value: Side | None) -> Sides:
        """
        Gets copy of instance with left, right, top, bottom sides set or removed

        Args:
            value (Side | None): Side value

        Returns:
            Sides: Sides instance
        """
        cp = self.copy()
        cp.prop_top = value
        cp.prop_bottom = value
        cp.prop_left = value
        cp.prop_right = value
        return cp

    def style_top(self, value: Side | None) -> Sides:
        """
        Gets a copy of instance with top side set or removed

        Args:
            value (Side | None): Side value

        Returns:
            Sides: Sides instance
        """
        cp = self.copy()
        cp.prop_top = value
        return cp

    def style_bottom(self, value: Side | None) -> Sides:
        """
        Gets a copy of instance with bottom side set or removed

        Args:
            value (Side | None): Side value

        Returns:
            Sides: Sides instance
        """
        cp = self.copy()
        cp.prop_bottom = value
        return cp

    def style_left(self, value: Side | None) -> Sides:
        """
        Gets a copy of instance with left side set or removed

        Args:
            value (Side | None): Side value

        Returns:
            Sides: Sides instance
        """
        cp = self.copy()
        cp.prop_left = value
        return cp

    def style_right(self, value: Side | None) -> Sides:
        """
        Gets a copy of instance with right side set or removed

        Args:
            value (Side | None): Side value

        Returns:
            Sides: Sides instance
        """
        cp = self.copy()
        cp.prop_right = value
        return cp

    # endregion style methods

    # region Properties
    @property
    def prop_style_kind(self) -> StyleKind:
        """Gets the kind of style"""
        return StyleKind.CHAR

    @property
    def prop_left(self) -> Side | None:
        """Gets left value"""
        return self._get("CharLeftBorder")

    @prop_left.setter
    def prop_left(self, value: Side | None) -> None:
        if value is None:
            self._remove("CharLeftBorder")
            return
        self._set("CharLeftBorder", value)

    @property
    def prop_right(self) -> Side | None:
        """Gets right value"""
        return self._get("CharRightBorder")

    @prop_right.setter
    def prop_right(self, value: Side | None) -> None:
        if value is None:
            self._remove("CharRightBorder")
            return
        self._set("CharRightBorder", value)

    @property
    def prop_top(self) -> Side | None:
        """Gets top value"""
        return self._get("CharTopBorder")

    @prop_top.setter
    def prop_top(self, value: Side | None) -> None:
        if value is None:
            self._remove("CharTopBorder")
            return
        self._set("CharTopBorder", value)

    @property
    def prop_bottom(self) -> Side | None:
        """Gets bottom value"""
        return self._get("CharBottomBorder")

    @prop_bottom.setter
    def prop_bottom(self, value: Side | None) -> None:
        if value is None:
            self._remove("CharBottomBorder")
            return
        self._set("CharBottomBorder", value)

    # endregion Properties
