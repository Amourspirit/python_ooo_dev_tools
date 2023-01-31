"""
Module for managing character font.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Any, Tuple, overload

from .....exceptions import ex as mEx
from .....utils import lo as mLo
from ....kind.format_kind import FormatKind
from ....style_base import StyleBase
from ...common.font_only_props import FontOnlyProps


class FontOnly(StyleBase):
    """
    Character Font

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together font properties.

    .. versionadded:: 0.9.0
    """

    def __init__(self, name: str, size: float, style_name: str) -> None:
        """
        Font options used in styles.

        Args:
            name (str): This property specifies the name of the font style. It may contain more than one name separated by comma.
            size (float): This value contains the size of the characters in point units.
            style_name (str): Font style name such as ``Bold``.
        """
        init_vals = {}
        init_vals[self._props.name] = name
        init_vals[self._props.style_name] = style_name
        init_vals[self._props.size] = size

        super().__init__(**init_vals)

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        return ("com.sun.star.style.CharacterProperties",)

    # region apply()
    @overload
    def apply(self, obj: object) -> None:
        ...

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies styles to object

        Args:
            obj (object): UNO object that has supports ``com.sun.star.style.CharacterProperties`` service.

        Returns:
            None:
        """
        super().apply(obj, **kwargs)

    def _props_set(self, obj: object, **kwargs: Any) -> None:
        try:
            super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"FontOnly.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion apply()
    # endregion methods

    # region Format Methods

    def fmt_size(self, value: float) -> FontOnly:
        """
        Get copy of instance with text size set.

        Args:
            value (float, optional): The size of the characters in point units of the font.

        Returns:
            FontOnly: Font with style added
        """
        ft = self.copy()
        ft.prop_size = value
        return ft

    def fmt_name(self, value: str) -> FontOnly:
        """
        Get copy of instance with name set.

        Args:
            value (str, optional): The name of the font.

        Returns:
            FontOnly: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_name = value
        return ft

    def fmt_style_name(self, value: str) -> FontOnly:
        """
        Get copy of instance with style name set.

        Args:
            value (str, optional): The style name of the font.

        Returns:
            FontOnly: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_style_name = value
        return ft

    # endregion Format Methods

    # region Style Properties

    # endregion Style Properties

    # region Prop Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.CHAR

    @property
    def prop_size(self) -> float:
        """This value contains the size of the characters in point."""
        return self._get(self._props.size)

    @prop_size.setter
    def prop_size(self, value: float) -> None:
        self._set(self._props.size, value)

    @property
    def prop_name(self) -> str:
        """This property specifies the name of the font style."""
        return self._get(self._props.name)

    @prop_name.setter
    def prop_name(self, value: str) -> None:
        self._set(self._props.name, value)

    @property
    def prop_style_name(self) -> str:
        """This property specifies the style name of the font style."""
        return self._get(self._props.style_name)

    @prop_style_name.setter
    def prop_style_name(self, value: str) -> None:
        self._set(self._props.style_name, value)

    @property
    def _props(self) -> FontOnlyProps:
        try:
            return self._font_only_properties
        except AttributeError:
            self._font_only_properties = FontOnlyProps(
                name="CharFontName", size="CharHeight", style_name="CharFontStyleName"
            )
        return self._font_only_properties

    # endregion Prop Properties
