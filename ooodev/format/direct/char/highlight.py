"""
Module for handeling character highlight.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple, overload

from ....exceptions import ex as mEx
from ....meta.static_prop import static_prop
from ....utils import lo as mLo
from ....utils import props as mProps
from ....utils.color import Color
from ...kind.style_kind import StyleKind
from ...style_base import StyleBase


class Highlight(StyleBase):
    """
    Character Highlighting

    .. versionadded:: 0.9.0
    """

    _EMPTY = None

    def __init__(self, color: Color = -1) -> None:
        """
        Constructor

        Args:
            color (Color, optional): Highlight Color

        Returns:
            None:
        """
        init_vals = {}
        if color >= 0:
            init_vals["CharBackColor"] = color
            init_vals["CharBackTransparent"] = False
        else:
            init_vals["CharBackColor"] = -1
            init_vals["CharBackTransparent"] = True

        super().__init__(**init_vals)

    def _supported_services(self) -> Tuple[str, ...]:
        """
        Gets a tuple of supported services (``com.sun.star.style.CharacterProperties``,)

        Returns:
            Tuple[str, ...]: Supported services
        """
        return ("com.sun.star.style.CharacterProperties",)

    # region apply_style()

    @overload
    def apply(self, obj: object) -> None:
        ...

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies padding to ``obj``

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.CharacterProperties`` service.
            kwargs (Any, optional): Expandable list of key value pairs that may be used in child classes.

        Returns:
            None:
        """
        try:
            super().apply(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__}.apply_style(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")
        return None

    # endregion apply_style()

    @staticmethod
    def from_obj(obj: object) -> Highlight:
        """
        Gets Hightlight instance from object

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.CharacterProperties`` service.

        Raises:
            NotSupportedServiceError: If ``obj`` does not support ``com.sun.star.style.CharacterProperties`` service.

        Returns:
            Hightlight: Hightlight that represents ``obj`` Hightlight.
        """
        inst = Highlight()
        if inst._is_valid_service(obj):
            inst._set("CharBackColor", int(mProps.Props.get(obj, "CharBackColor")))
            inst._set("CharBackTransparent", bool(mProps.Props.get(obj, "CharBackTransparent")))
        else:
            raise mEx.NotSupportedServiceError(inst._supported_services()[0])
        return inst

    # region set styles
    def style_color(self, value: int) -> Highlight:
        """
        Gets copy of instance with color set.

        Args:
            value (float | None): color value. If value is less then zero then it means no color.

        Returns:
            Highlight: Highlight instance
        """
        cp = self.copy()
        if value < 0:
            cp.prop_color = -1
        else:
            cp.prop_color = value

    # endregion set styles
    @property
    def prop_style_kind(self) -> StyleKind:
        """Gets the kind of style"""
        return StyleKind.CHAR

    @property
    def prop_color(self) -> int:
        """Gets/Sets color"""
        return self._get("CharBackColor")

    @prop_color.setter
    def prop_color(self, value: int):
        if value >= 0:
            self._set("CharBackColor", value)
            self._set("CharBackTransparent", False)
        else:
            self._set("CharBackColor", -1)
            self._set("CharBackTransparent", True)

    @static_prop
    def empty() -> Highlight:  # type: ignore[misc]
        """Gets Highlight empty. Static Property."""
        if Highlight._EMPTY is None:
            Highlight._EMPTY = Highlight()
        return Highlight._EMPTY
