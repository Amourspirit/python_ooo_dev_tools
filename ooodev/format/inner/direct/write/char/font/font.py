"""
Module for managing character fonts.

.. versionadded:: 0.9.0
"""

# region Import
from __future__ import annotations
from typing import Any, Tuple, cast, overload, TypeVar
import contextlib
import uno
from ooo.dyn.awt.char_set import CharSetEnum
from ooo.dyn.awt.font_family import FontFamilyEnum
from ooo.dyn.awt.font_slant import FontSlant
from ooo.dyn.awt.font_strikeout import FontStrikeoutEnum
from ooo.dyn.awt.font_underline import FontUnderlineEnum
from ooo.dyn.awt.font_weight import FontWeightEnum
from ooo.dyn.table.shadow_format import ShadowFormat

from ooodev.exceptions import ex as mEx
from ooodev.format.inner.direct.write.char.font.font_effects import FontLine
from ooodev.format.inner.direct.write.char.font.font_position import CharSpacingKind
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.loader import lo as mLo
from ooodev.units.angle import Angle
from ooodev.units.unit_convert import UnitConvert
from ooodev.units.unit_obj import UnitT
from ooodev.units.unit_pt import UnitPT
from ooodev.utils import info as mInfo
from ooodev.utils.color import Color

# endregion Import

_TFont = TypeVar("_TFont", bound="Font")


class Font(StyleBase):
    """
    Character Font

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together font properties.

    Many properties such as ``bold``, ``italic``, ``underline`` can be chained together.

    Example:

        .. code-block:: python

            # chaining fonts together to add new properties
            ft = Font().bold.italic.underline

            ft_color = Font().style_color(CommonColor.GREEN).style_bg_color(CommonColor.LIGHT_BLUE)

    .. seealso::

        - :ref:`help_writer_format_direct_char_font`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        b: bool | None = None,
        i: bool | None = None,
        u: bool | None = None,
        bg_color: Color | None = None,
        bg_transparent: bool | None = None,
        charset: CharSetEnum | None = None,
        color: Color | None = None,
        family: FontFamilyEnum | None = None,
        name: str | None = None,
        overline: FontLine | None = None,
        rotation: int | Angle | None = None,
        shadow_fmt: ShadowFormat | None = None,
        shadowed: bool | None = None,
        size: float | UnitT | None = None,
        slant: FontSlant | None = None,
        spacing: CharSpacingKind | float | UnitT | None = None,
        strike: FontStrikeoutEnum | None = None,
        subscript: bool | None = None,
        superscript: bool | None = None,
        underline: FontLine | None = None,
        weight: FontWeightEnum | float | None = None,
        word_mode: bool | None = None,
    ) -> None:
        """
        Font options used in styles.

        Args:
            b (bool, optional): Shortcut to set ``weight`` to bold.
            i (bool, optional): Shortcut to set ``slant`` to italic.
            u (bool, optional): Shortcut ot set ``underline`` to underline.
            bg_color (:py:data:`~.utils.color.Color`, optional): The value of the text background color.
            bg_transparent (bool, optional): Determines if the text background color is set to transparent.
            charset (CharSetEnum, optional): The text encoding of the font.
            color (:py:data:`~.utils.color.Color`, optional): The value of the text color. Setting to ``-1`` will cause automatic color.
            family (FontFamilyEnum, optional): Font Family.
            name (str, optional): This property specifies the name of the font style.
                It may contain more than one name separated by comma.
            overline (FontLine, optional): Character overline values.
            rotation (int, Angle, optional): Specifies the rotation of a character in degrees.
                Depending on the implementation only certain values may be allowed.
            shadow_fmt: (ShadowFormat, optional): Determines the type, color, and width of the shadow.
            shadowed (bool, optional): Specifies if the characters are formatted and displayed with a shadow effect.
            size (float, UnitT, optional): This value contains the size of the characters in ``pt`` (point) units
                or :ref:`proto_unit_obj`.
            slant (FontSlant, optional): The value of the posture of the document such as ``FontSlant.ITALIC``.
            spacing (CharSpacingKind, float, UnitT, optional): Specifies character spacing in ``pt`` (point) units
                or :ref:`proto_unit_obj`.
            strike (FontStrikeoutEnum, optional): Determines the type of the strike out of the character.
            subscript (bool, optional): Subscript option.
            superscript (bool, optional): Superscript option.
            underline (FontLine, optional): Character underline values.
            weight (FontWeightEnum, float, optional): The value of the font weight.
            word_mode(bool, optional): If ``True``, the underline and strike-through properties are not applied to
                white spaces.

        Returns:
            None:

        See Also:

            - :ref:`help_writer_format_direct_char_font`
        """
        # sourcery skip: low-code-quality
        # could not find any documentation in the API or elsewhere online for Overline
        # see: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1CharacterProperties.html

        super().__init__()
        if b is not None:
            self.prop_is_bold = b
        if i is not None:
            self.prop_is_italic = i
        if u is not None:
            self.prop_is_underline = u
        if bg_color is not None:
            self.prop_bg_color = bg_color
        if bg_transparent is not None:
            self.prop_bg_color_transparent = bg_transparent
        if charset is not None:
            self.prop_charset = charset
        if color is not None:
            self.prop_color = color
        if family is not None:
            self.prop_family = family
        if name is not None:
            self.prop_name = name
        if overline is not None:
            self.prop_overline = overline
        if rotation is not None:
            self.prop_rotation = rotation
        if shadow_fmt is not None:
            self.prop_shadow_fmt = shadow_fmt
        if shadowed is not None:
            self.prop_shadowed = shadowed
        if size is not None:
            self.prop_size = size
        if slant is not None:
            self.prop_slant = slant
        if spacing is not None:
            self.prop_spacing = spacing
        if strike is not None:
            self.prop_strike = strike
        # superscript and subscript use the same internal properties,CharEscapementHeight, CharEscapement
        if subscript is not None:
            self.prop_subscript = subscript
        if superscript is not None:
            self.prop_superscript = superscript
        if underline is not None:
            self.prop_underline = underline
        if weight is not None:
            self.prop_weight = weight
        if word_mode is not None:
            self.prop_word_mode = word_mode

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.style.CharacterProperties",
                "com.sun.star.style.CharacterStyle",
                "com.sun.star.style.ParagraphStyle",
                "com.sun.star.drawing.ControlShape",
            )
        return self._supported_services_values

    # region apply()
    @overload
    def apply(self, obj: Any) -> None:  # type: ignore
        ...

    @overload
    def apply(self, obj: Any, **kwargs) -> None:  # type: ignore
        ...

    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies styles to object

        Args:
            obj (object): UNO object that has supports ``com.sun.star.style.CharacterProperties`` service.

        Returns:
            None:
        """
        super().apply(obj, **kwargs)

    def _props_set(self, obj: Any, **kwargs: Any) -> None:
        try:
            super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print("Font.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion apply()
    # endregion methods

    # region Format Methods

    def fmt_bg_color(self: _TFont, value: Color | None = None) -> _TFont:
        """
        Get copy of instance with text background color set or removed.

        Args:
            value (:py:data:`~.utils.color.Color`, optional): The text background color.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_bg_color = value
        return ft

    def fmt_bg_transparent(self: _TFont, value: bool | None = None) -> _TFont:
        """
        Get copy of instance with text background transparency set or removed.

        Args:
            value (bool, optional): The text background transparency.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_bg_color_transparent = value
        return ft

    def fmt_charset(self: _TFont, value: CharSetEnum | None = None) -> _TFont:
        """
        Gets a copy of instance with charset set or removed.

        Args:
            value (CharSetEnum, optional): The text encoding of the font.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_charset = value
        return ft

    def fmt_color(self: _TFont, value: Color | None = None) -> _TFont:
        """
        Get copy of instance with text color set or removed.

        Args:
            value (:py:data:`~.utils.color.Color`, optional): The text color.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_color = value
        return ft

    def fmt_family(self: _TFont, value: FontFamilyEnum | None = None) -> _TFont:
        """
        Gets a copy of instance with charset set or removed.

        Args:
            value (FontFamilyEnum, optional): The Font Family.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_family = value
        return ft

    def fmt_size(self: _TFont, value: float | UnitT | None = None) -> _TFont:
        """
        Get copy of instance with text size set or removed.

        Args:
            value (float, UnitT, optional): The size of the characters in ``pt`` (point) units :ref:`proto_unit_obj`.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_size = value
        return ft

    def fmt_name(self: _TFont, value: str | None = None) -> _TFont:
        """
        Get copy of instance with name set or removed.

        Args:
            value (str, optional): The name of the font style. It may contain more than one name separated by comma.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_name = value
        return ft

    def fmt_overline(self: _TFont, value: FontUnderlineEnum | None = None) -> _TFont:
        """
        Get copy of instance with overline set or removed.

        Args:
            value (FontUnderlineEnum, optional): The size of the characters in point units.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ft = self.copy()
        fl = ft.prop_overline
        fl.line = value
        ft.prop_overline = fl
        return ft

    def fmt_overline_color(self: _TFont, value: Color | None = None) -> _TFont:
        """
        Get copy of instance with text overline color set or removed.

        Args:
            value (:py:data:`~.utils.color.Color`, optional): The color is used for an overline.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ft = self.copy()
        fl = ft.prop_overline
        fl.color = value
        ft.prop_overline = fl
        return ft

    def fmt_rotation(self: _TFont, value: int | Angle | None = None) -> _TFont:
        """
        Get copy of instance with rotation set or removed.

        Args:
            value (int, Angle, optional): The rotation of a character in degrees. Depending on the implementation only certain values may be allowed.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_rotation = value
        return ft

    def fmt_slant(self: _TFont, value: FontSlant | None = None) -> _TFont:
        """
        Get copy of instance with slant set or removed.

        Args:
            value (FontSlant, optional): The value of the posture of the document such as ``FontSlant.ITALIC``.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed

        Note:
            This method changes or removes any italic settings.
        """
        ft = self.copy()
        ft.prop_slant = value
        return ft

    def fmt_spacing(self: _TFont, value: float | UnitT | None = None) -> _TFont:
        """
        Get copy of instance with spacing set or removed.

        Args:
            value (float, UnitT, optional): The character spacing in ``pt`` (point) units :ref:`proto_unit_obj`.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_spacing = value
        return ft

    def fmt_shadow_fmt(self: _TFont, value: ShadowFormat | None = None) -> _TFont:
        """
        Get copy of instance with shadow format set or removed.

        Args:
            value (ShadowFormat, optional): The type, color, and width of the shadow.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_shadow_fmt = value
        return ft

    def fmt_strike(self: _TFont, value: FontStrikeoutEnum | None = None) -> _TFont:
        """
        Get copy of instance with strike set or removed.

        Args:
            value (FontStrikeoutEnum, optional): The type of the strike out of the character.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_strike = value
        return ft

    def fmt_subscript(self: _TFont, value: bool | None = None) -> _TFont:
        """
        Get copy of instance with text subscript set or removed.

        Args:
            value (bool, optional): The subscript.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_subscript = value
        return ft

    def fmt_superscript(self: _TFont, value: bool | None = None) -> _TFont:
        """
        Get copy of instance with text superscript set or removed.

        Args:
            value (bool, optional): The superscript.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_superscript = value
        return ft

    def fmt_underline(self: _TFont, value: FontUnderlineEnum | None = None) -> _TFont:
        """
        Get copy of instance with underline set or removed.

        Args:
            value (FontUnderlineEnum, optional): The value for the character underline.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ft = self.copy()
        fl = ft.prop_underline
        fl.line = value
        ft.prop_underline = fl
        return ft

    def fmt_underline_color(self: _TFont, value: Color | None = None) -> _TFont:
        """
        Gets copy of instance with text underline color set or removed.

        Args:
            value (:py:data:`~.utils.color.Color`, optional): The color is used for an underline.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ft = self.copy()
        fl = ft.prop_underline
        fl.color = value
        ft.prop_underline = fl
        return ft

    def fmt_weight(self: _TFont, value: FontWeightEnum | None = None) -> _TFont:
        """
        Get copy of instance with weight set or removed.

        Args:
            value (FontWeightEnum, optional): The value of the font weight.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed

        Note:
            This method changes or removes any bold settings.
        """
        ft = self.copy()
        ft.prop_weight = value
        return ft

    def fmt_word_mode(self: _TFont, value: bool | None = None) -> _TFont:
        """
        Get copy of instance with word mode set or removed.

        The underline and strike-through properties are not applied to white spaces when set to ``True``.

        Args:
            value (bool, optional): The word mode.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_word_mode = value
        return ft

    # endregion Format Methods

    # region Style Properties

    @property
    def bold(self: _TFont) -> _TFont:
        """Gets copy of instance with bold set"""
        # ft = self.copy()
        ft = self.copy()
        ft.prop_is_bold = True
        return ft

    @property
    def italic(self: _TFont) -> _TFont:
        """Gets copy of instance with italic set"""
        ft = self.copy()
        ft.prop_is_italic = True
        return ft

    @property
    def underline(self: _TFont) -> _TFont:
        """Gets copy of instance with underline set"""
        ft = self.copy()
        ft.prop_is_underline = True
        return ft

    @property
    def bg_transparent(self: _TFont) -> _TFont:
        """Gets copy of instance with background transparent set"""
        ft = self.copy()
        ft.prop_bg_color_transparent = True
        return ft

    @property
    def overline(self: _TFont) -> _TFont:
        """Gets copy of instance with overline set"""
        ft = self.copy()
        fl = self.prop_overline
        fl.line = FontUnderlineEnum.SINGLE
        ft.prop_overline = fl
        return ft

    @property
    def shadowed(self: _TFont) -> _TFont:
        """Gets copy of instance with shadow set"""
        ft = self.copy()
        ft.prop_shadowed = True
        return ft

    @property
    def strike(self: _TFont) -> _TFont:
        """Gets copy of instance with strike set"""
        ft = self.copy()
        ft.prop_strike = FontStrikeoutEnum.SINGLE
        return ft

    @property
    def subscript(self: _TFont) -> _TFont:
        """Gets copy of instance with sub script set"""
        ft = self.copy()
        ft.prop_subscript = True
        return ft

    @property
    def superscript(self: _TFont) -> _TFont:
        """Gets copy of instance with super script set"""
        ft = self.copy()
        ft.prop_superscript = True
        return ft

    @property
    def word_mode(self: _TFont) -> _TFont:
        """Gets copy of instance with word mode set"""
        ft = self.copy()
        ft.prop_word_mode = True
        return ft

    @property
    def spacing_very_tight(self: _TFont) -> _TFont:
        """Gets copy of instance with spacing set to very tight value"""
        ft = self.copy()
        ft.prop_spacing = CharSpacingKind.VERY_TIGHT
        return ft

    @property
    def spacing_tight(self: _TFont) -> _TFont:
        """Gets copy of instance with spacing set to tight value"""
        ft = self.copy()
        ft.prop_spacing = CharSpacingKind.TIGHT
        return ft

    @property
    def spacing_normal(self: _TFont) -> _TFont:
        """Gets copy of instance with spacing set to normal value"""
        ft = self.copy()
        ft.prop_spacing = CharSpacingKind.NORMAL
        return ft

    @property
    def spacing_loose(self: _TFont) -> _TFont:
        """Gets copy of instance with spacing set to loose value"""
        ft = self.copy()
        ft.prop_spacing = CharSpacingKind.LOOSE
        return ft

    @property
    def spacing_very_loose(self: _TFont) -> _TFont:
        """Gets copy of instance with spacing set to very loose value"""
        ft = self.copy()
        ft.prop_spacing = CharSpacingKind.VERY_LOOSE
        return ft

    # endregion Style Properties

    # region Prop Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.CHAR
        return self._format_kind_prop

    @property
    def prop_is_bold(self) -> bool:
        """Specifies bold"""
        pv = cast(float, self._get("CharWeight"))
        return pv == FontWeightEnum.BOLD.value if pv is not None else False

    @prop_is_bold.setter
    def prop_is_bold(self, value: bool | None) -> None:
        if value is None:
            self._remove("CharWeight")
            return
        if value:
            self._set("CharWeight", FontWeightEnum.BOLD.value)
        else:
            self._set("CharWeight", FontWeightEnum.NORMAL.value)

    @property
    def prop_bg_color(self) -> Color | None:
        """This property contains the text background color."""
        return self._get("CharBackColor")

    @prop_bg_color.setter
    def prop_bg_color(self, value: Color | None) -> None:
        if value is None:
            self._remove("CharBackColor")
            return
        self._set("CharBackColor", value)

    @property
    def prop_bg_color_transparent(self) -> bool | None:
        """This property contains the text background color."""
        return self._get("CharBackTransparent")

    @prop_bg_color_transparent.setter
    def prop_bg_color_transparent(self, value: bool | None) -> None:
        if value is None:
            self._remove("CharBackTransparent")
            return
        self._set("CharBackTransparent", value)

    @property
    def prop_is_italic(self) -> bool | None:
        """Specifies italic"""
        pv = cast(FontSlant, self._get("CharPosture"))
        return pv == FontSlant.ITALIC if pv is not None else None

    @prop_is_italic.setter
    def prop_is_italic(self, value: bool | None) -> None:
        if value is None:
            self._remove("CharPosture")
            return
        if value:
            self._set("CharPosture", FontSlant.ITALIC)
        else:
            self._set("CharPosture", FontSlant.NONE)

    @property
    def prop_is_underline(self) -> bool | None:
        """Specifies underline"""
        pv = cast(int, self._get("CharUnderline"))
        return pv != FontUnderlineEnum.NONE.value if pv is not None else None

    @prop_is_underline.setter
    def prop_is_underline(self, value: bool | None) -> None:
        if value is None:
            self._remove("CharUnderline")
            return
        if value:
            self._set("CharUnderline", FontUnderlineEnum.SINGLE.value)
        else:
            self._set("CharUnderline", FontUnderlineEnum.NONE.value)

    @property
    def prop_underline(self) -> FontLine:
        """This property contains the value for the character underline."""
        pv = cast(int, self._get("CharUnderline"))
        line = None if pv is None else FontUnderlineEnum(pv)
        return FontLine(line=line, color=cast(Color, self._get("CharUnderlineColor")))

    @prop_underline.setter
    def prop_underline(self, value: FontLine | None) -> None:
        if value is None:
            self._remove("CharUnderline")
            self._remove("CharUnderlineColor")
            self._remove("CharUnderlineHasColor")
            return
        if value.line is None:
            self._remove("CharUnderline")
        else:
            self._set("CharUnderline", value.line.value)

        if value.color is None:
            self._remove("CharUnderlineColor")
            self._remove("CharUnderlineHasColor")
        elif value.color < 0:
            # automatic color
            self._set("CharUnderlineHasColor", False)
            self._set("CharUnderlineColor", -1)
        else:
            self._set("CharUnderlineHasColor", True)
            self._set("CharUnderlineColor", value.color)

    @property
    def prop_charset(self) -> CharSetEnum | None:
        """This property contains the text encoding of the font."""
        pv = cast(int, self._get("CharFontCharSet"))
        return CharSetEnum(pv) if pv is not None else None

    @prop_charset.setter
    def prop_charset(self, value: CharSetEnum | None) -> None:
        if value is None:
            self._remove("CharFontCharSet")
            return
        self._set("CharFontCharSet", value.value)

    @property
    def prop_color(self) -> Color | None:
        """This property contains the value of the text color."""
        return self._get("CharColor")

    @prop_color.setter
    def prop_color(self, value: Color | None) -> None:
        if value is None:
            self._remove("CharColor")
            return
        self._set("CharColor", value)

    @property
    def prop_family(self) -> FontFamilyEnum | None:
        """This property contains font family."""
        pv = cast(FontFamilyEnum, self._get("CharFontFamily"))
        return FontFamilyEnum(pv) if pv is not None else None

    @prop_family.setter
    def prop_family(self, value: FontFamilyEnum | None) -> None:
        if value is None:
            self._remove("CharFontFamily")
            return
        self._set("CharFontFamily", value.value)

    @property
    def prop_size(self) -> UnitPT | None:
        """This value contains the size of the characters in ``pt`` (point) units."""
        pv = cast(float, self._get("CharHeight"))
        return None if pv is None else UnitPT(pv)

    @prop_size.setter
    def prop_size(self, value: float | UnitT | None) -> None:
        if value is None:
            self._remove("CharHeight")
            return
        try:
            self._set("CharHeight", value.get_value_pt())  # type: ignore
        except AttributeError:
            self._set("CharHeight", float(value))  # type: ignore

    @property
    def prop_name(self) -> str | None:
        """This property specifies the name of the font style. It may contain more than one name separated by comma."""
        return self._get("CharFontName")

    @prop_name.setter
    def prop_name(self, value: str | None) -> None:
        if value is None:
            self._remove("CharFontName")
            return
        self._set("CharFontName", value)

    @property
    def prop_strike(self) -> FontStrikeoutEnum | None:
        """This property determines the type of the strike out of the character."""
        pv = cast(int, self._get("CharStrikeout"))
        return FontStrikeoutEnum(pv) if pv is not None else None

    @prop_strike.setter
    def prop_strike(self, value: FontStrikeoutEnum | None) -> None:
        if value is None:
            self._remove("CharStrikeout")
            return
        self._set("CharStrikeout", value.value)

    @property
    def prop_weight(self) -> FontWeightEnum | None:
        """This property contains the value of the font weight."""
        pv = cast(float, self._get("CharWeight"))
        return FontWeightEnum(pv) if pv is not None else None

    @prop_weight.setter
    def prop_weight(self, value: FontWeightEnum | float | None) -> None:
        if value is None:
            self._remove("CharWeight")
            return
        with contextlib.suppress(AttributeError):
            self._set("CharWeight", value.value)  # type: ignore
            return
        self._set("CharWeight", float(value))  # type: ignore

    @property
    def prop_slant(self) -> FontSlant | None:
        """This property contains the value of the posture of the document such as  ``FontSlant.ITALIC``"""
        return self._get("CharPosture")

    @prop_slant.setter
    def prop_slant(self, value: FontSlant | None) -> None:
        if value is None:
            self._remove("CharPosture")
            return
        self._set("CharPosture", value)

    @property
    def prop_spacing(self) -> UnitPT | None:
        """This value contains character spacing in ``pt`` (point) units"""
        pv = self._get("CharKerning")
        return None if pv is None else UnitPT.from_mm100(pv)

    @prop_spacing.setter
    def prop_spacing(self, value: float | CharSpacingKind | UnitT | None) -> None:
        if value is None:
            self._remove("CharKerning")
            return
        try:
            self._set("CharKerning", value.get_value_mm100())  # type: ignore
        except AttributeError:
            self._set("CharKerning", UnitConvert.convert_pt_mm100(float(value)))  # type: ignore

    @property
    def prop_shadowed(self) -> bool | None:
        """This property specifies if the characters are formatted and displayed with a shadow effect."""
        return self._get("CharShadowed")

    @prop_shadowed.setter
    def prop_shadowed(self, value: bool | None) -> None:
        if value is None:
            self._remove("CharShadowed")
            return
        self._set("CharShadowed", value)

    @property
    def prop_shadow_fmt(self) -> ShadowFormat | None:
        """This property specifies the type, color, and width of the shadow."""
        return self._get("CharShadowFormat")

    @prop_shadow_fmt.setter
    def prop_shadow_fmt(self, value: ShadowFormat | None) -> None:
        if value is None:
            self._remove("CharShadowFormat")
            return
        if mInfo.Info.is_type_struct(value, "com.sun.star.table.ShadowFormat"):
            self._set("CharShadowFormat", value)

    @property
    def prop_superscript(self) -> bool | None:
        """Specifies if the font is super script"""
        pv = cast(int, self._get("CharEscapement"))
        return pv > 0 if pv is not None else None

    @prop_superscript.setter
    def prop_superscript(self, value: bool | None) -> None:
        if value is None:
            self._remove("CharEscapementHeight")
            self._remove("CharEscapement")
            return
        if value:
            self._set("CharEscapementHeight", 58)
            self._set("CharEscapement", 14_000)
        else:
            self._set("CharEscapementHeight", 100)
            self._set("CharEscapement", 0)

    @property
    def prop_subscript(self) -> bool | None:
        """Specifies if the font is sub script"""
        pv = cast(int, self._get("CharEscapement"))
        return pv < 0 if pv is not None else None

    @prop_subscript.setter
    def prop_subscript(self, value: bool | None) -> None:
        if value is None:
            self._remove("CharEscapementHeight")
            self._remove("CharEscapement")
            return
        if value:
            self._set("CharEscapementHeight", 58)
            self._set("CharEscapement", -14_000)
        else:
            self._set("CharEscapementHeight", 100)
            self._set("CharEscapement", 0)

    @property
    def prop_overline(self) -> FontLine:
        """This property contains the value for the character overline."""
        pv = cast(int, self._get("CharOverline"))
        line = None if pv is None else FontUnderlineEnum(pv)
        return FontLine(line=line, color=cast(Color, self._get("CharOverlineColor")))

    @prop_overline.setter
    def prop_overline(self, value: FontLine | None) -> None:
        if value is None:
            self._remove("CharOverline")
            self._remove("CharOverlineColor")
            self._remove("CharOverlineHasColor")
            return
        if value.line is None:
            self._remove("CharOverline")
        else:
            self._set("CharOverline", value.line.value)

        if value.color is None:
            self._remove("CharOverlineColor")
            self._remove("CharOverlineHasColor")
        elif value.color < 0:
            # automatic color
            self._set("CharOverlineHasColor", False)
            self._set("CharOverlineColor", -1)
        else:
            self._set("CharOverlineHasColor", True)
            self._set("CharOverlineColor", value.color)

    @property
    def prop_rotation(self) -> Angle | None:
        """
        This optional property determines the rotation of a character in degrees.

        Depending on the implementation only certain values may be allowed.
        """
        pv = cast(int, self._get("CharRotation"))
        return None if pv is None else Angle(round(pv / 10))

    @prop_rotation.setter
    def prop_rotation(self, value: int | Angle | None) -> None:
        if value is None:
            self._remove("CharRotation")
            return
        self._set("CharRotation", Angle(int(value)).value * 10)

    @property
    def prop_word_mode(self) -> bool | None:
        """If this property is ``True``, the underline and strike-through properties are not applied to white spaces."""
        return self._get("CharWordMode")

    @prop_word_mode.setter
    def prop_word_mode(self, value: bool | None) -> None:
        if value is None:
            self._remove("CharWordMode")
            return
        self._set("CharWordMode", value)

    # endregion Prop Properties
