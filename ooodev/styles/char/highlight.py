"""
Module for handeling character highlight.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import overload

from ...exceptions import ex as mEx
from ...meta.static_prop import static_prop
from ...utils import info as mInfo
from ...utils import lo as mLo
from ...utils import props as mProps
from ..style_base import StyleBase
from ...utils.color import Color


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

    def _is_supported(self, obj: object) -> bool:
        return mInfo.Info.support_service(obj, "com.sun.star.style.CharacterProperties")

    # region apply_style()

    @overload
    def apply_style(self, obj: object) -> None:
        ...

    def apply_style(self, obj: object, **kwargs) -> None:
        """
        Applies padding to ``obj``

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.CharacterProperties`` service.
            kwargs (Any, optional): Expandable list of key value pairs that may be used in child classes.

        Returns:
            None:
        """
        if mInfo.Info.support_service(obj, "com.sun.star.style.CharacterProperties"):
            try:
                super().apply_style(obj)
            except mEx.MultiError as e:
                mLo.Lo.print(f"{self.__class__}.apply_style(): Unable to set Property")
                for err in e.errors:
                    mLo.Lo.print(f"  {err}")
        else:
            mLo.Lo.print(f'{self.__class__}.apply_style(): "com.sun.star.style.CharacterProperties" not supported')
        return None

    # endregion apply_style()

    @staticmethod
    def from_obj(obj: object) -> Highlight:
        """
        Gets Hightlight instance from object

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.CharacterProperties`` service.

        Raises:
            NotSupportedServiceError: If ``obj`` does not support  ``com.sun.star.style.CharacterProperties`` service.

        Returns:
            Hightlight: Hightlight that represents ``obj`` Hightlight.
        """
        if not mInfo.Info.support_service(obj, "com.sun.star.style.CharacterProperties"):
            raise mEx.NotSupportedServiceError("com.sun.star.style.CharacterProperties")
        inst = Highlight()
        inst._set("CharBackColor", int(mProps.Props.get(obj, "CharBackColor")))
        inst._set("CharBackTransparent", bool(mProps.Props.get(obj, "CharBackTransparent")))
        return inst

    @property
    def color(self) -> int:
        """Gets/Sets color"""
        return self._get("CharBackColor")

    @color.setter
    def color(self, value: int):
        if value >= 0:
            self._set("CharBackColor", value)
            self._set("CharBackTransparent", False)
        else:
            self._set("CharBackColor", -1)
            self._set("CharBackTransparent", True)

    @static_prop
    def empty(cls) -> Highlight:
        """Gets Highlight empty. Static Property."""
        if cls._EMPTY is None:
            cls._EMPTY = Highlight()
        return cls._EMPTY
