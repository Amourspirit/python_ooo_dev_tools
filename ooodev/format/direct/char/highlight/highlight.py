"""
Module for handeling character highlight.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple, Type, overload, TypeVar

from .....events.args.cancel_event_args import CancelEventArgs
from .....exceptions import ex as mEx
from .....meta.static_prop import static_prop
from .....utils import lo as mLo
from .....utils import props as mProps
from .....utils.color import Color
from ....kind.format_kind import FormatKind
from ....style_base import StyleBase

_THighlight = TypeVar(name="_THighlight", bound="Highlight")


class Highlight(StyleBase):
    """
    Character Highlighting

    .. versionadded:: 0.9.0
    """

    def __init__(self, color: Color = -1) -> None:
        """
        Constructor

        Args:
            color (Color, optional): Highlight Color. A value of ``-1`` Set color to Transparent.

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
        return (
            "com.sun.star.style.CharacterProperties",
            "com.sun.star.style.CharacterStyle",
            "com.sun.star.style.ParagraphStyle",
        )

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
            obj (object): UNO object that supports ``com.sun.star.style.CharacterProperties`` service.

        Returns:
            None:
        """
        try:
            super().apply(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")
        return None

    # endregion apply()

    @classmethod
    def from_obj(cls: Type[_THighlight], obj: object) -> _THighlight:
        """
        Gets Hightlight instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Hightlight: Hightlight that represents ``obj`` Hightlight.
        """
        inst = super(Highlight, cls).__new__(cls)
        inst.__init__()
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        inst._set("CharBackColor", int(mProps.Props.get(obj, "CharBackColor")))
        inst._set("CharBackTransparent", bool(mProps.Props.get(obj, "CharBackTransparent")))

        return inst

    # region set styles
    def fmt_color(self: _THighlight, value: Color) -> _THighlight:
        """
        Gets copy of instance with color set.

        Args:
            value (int): color value. If value is less then zero then it means no color.

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
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.CHAR

    @property
    def prop_color(self) -> Color:
        """Gets/Sets color"""
        return self._get("CharBackColor")

    @prop_color.setter
    def prop_color(self, value: Color):
        if value >= 0:
            self._set("CharBackColor", value)
            self._set("CharBackTransparent", False)
        else:
            self._set("CharBackColor", -1)
            self._set("CharBackTransparent", True)

    @static_prop
    def empty() -> Highlight:  # type: ignore[misc]
        """Gets Highlight empty. Static Property."""
        try:
            return Highlight._EMPTY_INST
        except AttributeError:
            Highlight._EMPTY_INST = Highlight()
            Highlight._EMPTY_INST._is_default_inst = True
        return Highlight._EMPTY_INST