"""
Module for handling character highlight.

.. seealso::

    :ref:`help_writer_format_direct_char_highlight`

.. versionadded:: 0.9.0
"""

# region Import
from __future__ import annotations
from typing import Any, Tuple, Type, overload, TypeVar

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.utils.color import Color, StandardColor

# endregion Import
_THighlight = TypeVar("_THighlight", bound="Highlight")


class Highlight(StyleBase):
    """
    Character Highlighting

    .. versionadded:: 0.9.0
    """

    def __init__(self, color: Color = StandardColor.AUTO_COLOR) -> None:
        """
        Constructor

        Args:
            color (~ooodev.utils.color.Color, optional): Highlight Color. A value of ``-1`` Set color to Transparent.

        Returns:
            None:

        See Also:

            - :ref:`help_writer_format_direct_char_highlight`
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
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.style.CharacterProperties",
                "com.sun.star.style.CharacterStyle",
                "com.sun.star.style.ParagraphStyle",
            )
        return self._supported_services_values

    def _on_modifying(self, source: Any, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifying(source, event)

    # region apply()

    @overload
    def apply(self, obj: Any) -> None:  # type: ignore
        ...

    def apply(self, obj: Any, **kwargs) -> None:
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

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_THighlight], obj: Any) -> _THighlight: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_THighlight], obj: Any, **kwargs) -> _THighlight: ...

    @classmethod
    def from_obj(cls: Type[_THighlight], obj: Any, **kwargs) -> _THighlight:
        """
        Gets Highlight instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Highlight: Highlight that represents ``obj`` Highlight.
        """
        # pylint: disable=protected-access
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        inst._set("CharBackColor", int(mProps.Props.get(obj, "CharBackColor")))
        inst._set("CharBackTransparent", bool(mProps.Props.get(obj, "CharBackTransparent")))

        return inst

    # endregion from_obj()

    # region set styles
    def fmt_color(self: _THighlight, value: Color) -> _THighlight:
        """
        Gets copy of instance with color set.

        Args:
            value (~ooodev.utils.color.Color): color value. If value is less than zero it means no color.

        Returns:
            Highlight: Highlight instance
        """
        cp = self.copy()
        cp.prop_color = StandardColor.AUTO_COLOR if value < 0 else value
        return cp

    # endregion set styles

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.CHAR
        return self._format_kind_prop

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

    @property
    def empty(self: _THighlight) -> _THighlight:  # type: ignore[misc]
        """Gets Highlight empty."""
        # pylint: disable=protected-access
        # pylint: disable=unexpected-keyword-arg
        try:
            return self._empty_inst
        except AttributeError:
            self._empty_inst = self.__class__(_cattribs=self._get_internal_cattribs())
            self._empty_inst._is_default_inst = True
        return self._empty_inst
