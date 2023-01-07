"""
Module for managing character fonts.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import cast, overload
from enum import Enum

from ...exceptions import ex as mEx
from ...utils import info as mInfo
from ...utils import lo as mLo
from ...utils.color import Color
from ..style_base import StyleBase
from ..style_const import POINT_RATIO

from ooo.dyn.awt.char_set import CharSetEnum as CharSetEnum
from ooo.dyn.awt.font_family import FontFamilyEnum as FontFamilyEnum
from ooo.dyn.awt.font_slant import FontSlant as FontSlant
from ooo.dyn.awt.font_strikeout import FontStrikeoutEnum as FontStrikeoutEnum
from ooo.dyn.awt.font_underline import FontUnderlineEnum as FontUnderlineEnum
from ooo.dyn.awt.font_weight import FontWeightEnum as FontWeightEnum
from ooo.dyn.table.shadow_format import ShadowFormat as ShadowFormat
from ooo.dyn.table.shadow_location import ShadowLocation as ShadowLocation


class CharSpacingKind(float, Enum):
    """
    Character Spacing

    .. versionadded:: 0.9.0
    """

    VERY_TIGHT = -3.0
    TIGHT = -1.5
    NORMAL = 0.0
    LOOSE = 3.0
    VERY_LOOSE = 6.0


class Font(StyleBase):
    """
    Character Font

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``style_`` can be used to chain together font properties.

    Many properties such as ``bold``, ``italic``, ``underline`` can be chained together.

    Example:
        .. code-block:: python

            # chaining fonts together to add new properties
            ft = Font().bold.italic.underline

            ft_color = Font().style_color(CommonColor.GREEN).style_bg_color(CommonColor.LIGHT_BLUE)

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
        size: float | None = None,
        name: str | None = None,
        overline: FontUnderlineEnum | None = None,
        overline_color: Color | None = None,
        rotation: float | None = None,
        slant: FontSlant | None = None,
        spacing: CharSpacingKind | float | None = None,
        shadowed: bool | None = None,
        shadow_fmt: ShadowFormat | None = None,
        strike: FontStrikeoutEnum | None = None,
        subscript: bool | None = None,
        superscript: bool | None = None,
        underine: FontUnderlineEnum | None = None,
        underine_color: Color | None = None,
        weight: FontWeightEnum | None = None,
        word_mode: bool | None = None,
    ) -> None:
        """
        Font options used in styles.

        Args:
            b (bool, optional): Short cut to set ``weight`` to bold.
            i (bool, optional): Short cut to set ``slant`` to italic.
            u (bool, optional): Short cut ot set ``underline`` to underline.
            bg_color (Color, optional): The value of the text background color.
            bg_transparent (bool, optional): Determines if the text background color is set to transparent.
            charset (CharSetEnum, optional): The text encoding of the font.
            color (Color, optional): The value of the text color.
            family (FontFamilyEnum, optional): Font Family
            size (float, optional): This value contains the size of the characters in point units.
            name (str, optional): This property specifies the name of the font style. It may contain more than one name separated by comma.
            overline (FontUnderlineEnum, optional): The value for the character overline.
            overline_color (Color, optional): Specifies if the property ``CharOverlinelineColor`` is used for an overline.
            rotation (float, optional): Determines the rotation of a character in degrees. Depending on the implementation only certain values may be allowed.
            slant (FontSlant, optional): The value of the posture of the document such as ``FontSlant.ITALIC``.
            spacing (float, optional): Character spacing in point units.
            shadowed (bool, optional): Specifies if the characters are formatted and displayed with a shadow effect.
            shadow_fmt: (ShadowFormat, optional): Determines the type, color, and width of the shadow.
            strike (FontStrikeoutEnum, optional): Detrmines the type of the strike out of the character.
            subscript (bool, optional): Sub script option.
            superscript (bool, optional): Super script option.
            underine (FontUnderlineEnum, optional): The value for the character underline.
            underine_color (Color, optional): Specifies if the property ``CharUnderlineColor`` is used for an underline.
            weight (FontWeightEnum, optional): The value of the font weight.
            word_mode(bool, optional): If ``True``, the underline and strike-through properties are not applied to white spaces.
        """
        # could not find any documention in the API or elsewhere online for Overline
        # see: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1CharacterProperties.html
        init_vals = {
            "CharFontName": name,
            "CharColor": color,
            "CharBackColor": bg_color,
            "CharUnderlineColor": underine_color,
            "CharOverlineColor": overline_color,
            "CharHeight": size,
            "CharBackTransparent": bg_transparent,
            "CharWordMode": word_mode,
            "CharShadowed": shadowed,
        }
        if not bg_color is None:
            init_vals["CharBackTransparent"] = False

        if not overline_color is None:
            init_vals["CharOverlineHasColor"] = True
        if not underine_color is None:
            init_vals["CharUnderlineHasColor"] = True
        if not charset is None:
            init_vals["CharFontCharSet"] = charset.value
        if not family is None:
            init_vals["CharFontFamily"] = family.value
        if not strike is None:
            init_vals["CharStrikeout"] = strike.value

        if not b is None:
            if b:
                init_vals["CharWeight"] = FontWeightEnum.BOLD.value
            else:
                init_vals["CharWeight"] = FontWeightEnum.NORMAL.value
        if not i is None:
            if i:
                init_vals["CharPosture"] = FontSlant.ITALIC
            else:
                init_vals["CharWeight"] = FontSlant.NONE
        if not u is None:
            if u:
                init_vals["CharUnderline"] = FontUnderlineEnum.SINGLE.value
            else:
                init_vals["CharUnderline"] = FontUnderlineEnum.NONE.value

        if not overline is None:
            init_vals["CharOverline"] = overline.value

        if not underine is None:
            init_vals["CharUnderline"] = underine.value

        if not weight is None:
            init_vals["CharWeight"] = weight.value

        if not slant is None:
            init_vals["CharPosture"] = slant

        if not spacing is None:
            init_vals["CharKerning"] = round(float(spacing) * POINT_RATIO)

        if not rotation is None:
            init_vals["CharRotation"] = round(rotation * 10)

        if not shadow_fmt is None:
            if mInfo.Info.is_type_struct(shadow_fmt, "com.sun.star.table.ShadowFormat"):
                init_vals["CharShadowFormat"] = shadow_fmt

        super().__init__(**init_vals)

        # superscript and subscript use the same internal properties,CharEscapementHeight, CharEscapement
        if not superscript is None:
            self.prop_superscript = superscript
        if not subscript is None:
            self.prop_subscript = subscript

    # region methods

    @overload
    def apply_style(self, obj: object) -> None:
        ...

    def apply_style(self, obj: object, **kwargs) -> None:
        """
        Applies styles to object

        Args:
            obj (object): UNO object that has supports ``com.sun.star.style.CharacterProperties`` service.

        Returns:
            None:
        """
        if mInfo.Info.support_service(obj, "com.sun.star.style.CharacterProperties"):
            try:
                super().apply_style(obj)
            except mEx.MultiError as e:
                mLo.Lo.print(f"Font.apply_style(): Unable to set Property")
                for err in e.errors:
                    mLo.Lo.print(f"  {err}")
        else:
            mLo.Lo.print('Font.apply_style(): "com.sun.star.style.CharacterProperties" not supported')

    # endregion methods

    # region Style Methods

    def style_bg_color(self, value: Color | None = None) -> Font:
        """
        Get copy of instance with text background color set or removed.

        Args:
            value (Color, optional): The text background color.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_bg_color = value
        return ft

    def style_bg_transparent(self, value: bool | None = None) -> Font:
        """
        Get copy of instance with text background transparency set or removed.

        Args:
            value (bool, optional): The text background transparency.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_bg_transparent = value
        return ft

    def style_charset(self, value: CharSetEnum | None = None) -> Font:
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

    def style_color(self, value: Color | None = None) -> Font:
        """
        Get copy of instance with text color set or removed.

        Args:
            value (Color, optional): The text color.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_color = value
        return ft

    def style_family(self, value: FontFamilyEnum | None = None) -> Font:
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

    def style_size(self, value: float | None = None) -> Font:
        """
        Get copy of instance with text size set or removed.

        Args:
            value (float, optional): The size of the characters in point units.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_size = value
        return ft

    def style_name(self, value: str | None = None) -> Font:
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

    def style_overline(self, value: FontUnderlineEnum | None = None) -> Font:
        """
        Get copy of instance with overline set or removed.

        Args:
            value (FontUnderlineEnum, optional): The size of the characters in point units.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_overline = value
        return ft

    def style_overline_color(self, value: Color | None = None) -> Font:
        """
        Get copy of instance with text overline color set or removed.

        Args:
            value (Color, optional): The color is used for an overline.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_overline_color = value
        return ft

    def style_rotation(self, value: float | None = None) -> Font:
        """
        Get copy of instance with rotation set or removed.

        Args:
            value (str, optional): The rotation of a character in degrees. Depending on the implementation only certain values may be allowed.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_rotation = value
        return ft

    def style_slant(self, value: FontSlant | None = None) -> Font:
        """
        Get copy of instance with slant set or removed.

        Args:
            value (FontSlant, optional): The value of the posture of the document such as ``FontSlant.ITALIC``.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed

        Note:
            This method chanages or removes any italic settings.
        """
        ft = self.copy()
        ft.prop_slant = value
        return ft

    def style_spacing(self, value: float | None = None) -> Font:
        """
        Get copy of instance with spacing set or removed.

        Args:
            value (str, optional): The character spacing in point units.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_spacing = value
        return ft

    def style_shadow_fmt(self, value: ShadowFormat | None = None) -> Font:
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

    def style_strike(self, value: FontStrikeoutEnum | None = None) -> Font:
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

    def style_subscript(self, value: bool | None = None) -> Font:
        """
        Get copy of instance with text sub script set or removed.

        Args:
            value (bool, optional): The sub script.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_subscript = value
        return ft

    def style_superscript(self, value: bool | None = None) -> Font:
        """
        Get copy of instance with text super script set or removed.

        Args:
            value (bool, optional): The super script.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_superscript = value
        return ft

    def style_underline(self, value: FontUnderlineEnum | None = None) -> Font:
        """
        Get copy of instance with underline set or removed.

        Args:
            value (FontUnderlineEnum, optional): The value for the character underline.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_underline = value
        return ft

    def style_weight(self, value: FontWeightEnum | None = None) -> Font:
        """
        Get copy of instance with weight set or removed or removed.

        Args:
            value (FontWeightEnum, optional): The value of the font weight.
                If ``None`` style is removed. Default ``None``

        Returns:
            Font: Font with style added or removed

        Note:
            This method chanages or removes any bold settings.
        """
        ft = self.copy()
        ft.prop_weight = value
        return ft

    def style_word_mode(self, value: bool | None = None) -> Font:
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

    # endregion Style Methods

    # region Style Properties

    @property
    def bold(self) -> Font:
        """Gets copy of instance with bold set"""
        # ft = self.copy()
        ft = self.copy()
        ft.prop_is_bold = True
        return ft

    @property
    def italic(self) -> Font:
        """Gets copy of instance with italic set"""
        ft = self.copy()
        ft.prop_is_italic = True
        return ft

    @property
    def underline(self) -> Font:
        """Gets copy of instance with underline set"""
        ft = self.copy()
        ft.prop_is_underline = True
        return ft

    @property
    def bg_transparent(self) -> Font:
        """Gets copy of instance with background transparent set"""
        ft = self.copy()
        ft.prop_bg_transparent = True
        return ft

    @property
    def overline(self) -> Font:
        """Gets copy of instance with overline set"""
        ft = self.copy()
        ft.prop_overline = FontUnderlineEnum.SINGLE
        return ft

    @property
    def shadowed(self) -> Font:
        """Gets copy of instance with shadow set"""
        ft = self.copy()
        ft.prop_shadowed = True
        return ft

    @property
    def strike(self) -> Font:
        """Gets copy of instance with strike set"""
        ft = self.copy()
        ft.prop_strike = FontStrikeoutEnum.SINGLE
        return ft

    @property
    def subscript(self) -> Font:
        """Gets copy of instance with sub script set"""
        ft = self.copy()
        ft.prop_subscript = True
        return ft

    @property
    def superscript(self) -> Font:
        """Gets copy of instance with super script set"""
        ft = self.copy()
        ft.prop_superscript = True
        return ft

    @property
    def word_mode(self) -> Font:
        """Gets copy of instance with word mode set"""
        ft = self.copy()
        ft.prop_word_mode = True
        return ft

    # endregion Style Properties

    # region Prop Properties
    @property
    def prop_is_bold(self) -> bool:
        """Specifies bold"""
        pv = cast(float, self._get("CharWeight"))
        if not pv is None:
            return pv == FontWeightEnum.BOLD.value
        return False

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
        if not pv is None:
            return pv == FontSlant.ITALIC
        return None

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
        if not pv is None:
            return pv != FontUnderlineEnum.NONE.value
        return None

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
    def prop_underline(self) -> FontUnderlineEnum | None:
        """Specifies underline"""
        pv = cast(int, self._get("CharUnderline"))
        if not pv is None:
            return FontUnderlineEnum(pv)
        return None

    @prop_underline.setter
    def prop_underline(self, value: FontUnderlineEnum | None) -> None:
        if value is None:
            self._remove("CharUnderline")
            return
        self._set("CharUnderline", value.value)

    @property
    def prop_charset(self) -> CharSetEnum | None:
        """This property contains the text encoding of the font."""
        pv = cast(int, self._get("CharFontCharSet"))
        if not pv is None:
            return CharSetEnum(pv)
        return None

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
        if not pv is None:
            return FontFamilyEnum(pv)
        return None

    @prop_family.setter
    def prop_family(self, value: FontFamilyEnum | None) -> None:
        if value is None:
            self._remove("CharFontFamily")
            return
        self._set("CharFontFamily", value.value)

    @property
    def prop_size(self) -> float | None:
        """This value contains the size of the characters in point."""
        return self._get("CharHeight")

    @prop_size.setter
    def prop_size(self, value: float | None) -> None:
        if value is None:
            self._remove("CharHeight")
            return
        self._set("CharHeight", value)

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
        if not pv is None:
            return FontStrikeoutEnum(pv)
        return None

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
        if not pv is None:
            return FontWeightEnum(pv)
        return None

    @prop_weight.setter
    def prop_weight(self, value: FontWeightEnum | None) -> None:
        if value is None:
            self._remove("CharWeight")
            return
        self._set("CharWeight", value.value)

    @property
    def prop_slant(self) -> FontSlant | None:
        """This property contains the value of the posture of the document such as  ``FontSlant.ITALIC``"""
        return self._get("CharPosture")

    @prop_slant.setter
    def prop_slant(self, value: FontSlant | None) -> None:
        if value is None:
            self._remove("CharPosture")
            return
        self._set("CharPosture", value.value)

    @property
    def prop_spacing(self) -> float | None:
        """This value contains character spacing in point units"""
        pv = self._get("CharKerning")
        if not pv is None:
            if pv == 0.0:
                return 0.0
            return pv / POINT_RATIO
        return None

    @prop_spacing.setter
    def prop_spacing(self, value: float | None) -> None:
        if value is None:
            self._remove("CharKerning")
            return
        self._set("CharKerning", round(float(value) * POINT_RATIO))

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
            self._set("CharShadowFormat", value.value)

    @property
    def prop_superscript(self) -> bool | None:
        """Specifies if the font is super script"""
        pv = cast(int, self._get("CharEscapement"))
        if not pv is None:
            return pv > 0
        return None

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
        if not pv is None:
            return pv < 0
        return None

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
    def prop_overline(self) -> FontUnderlineEnum | None:
        """This property contains the value for the character overline."""
        pv = cast(int, self._get("CharOverline"))
        if not pv is None:
            return FontUnderlineEnum(pv)
        return None

    @prop_overline.setter
    def prop_overline(self, value: FontUnderlineEnum | None) -> None:
        if value is None:
            self._remove("CharOverline")
            return
        self._set("CharOverline", value.value)

    @property
    def prop_overline_color(self) -> Color | None:
        """This property specifies if the property ``CharOverlineColor`` is used for an overline."""
        return self._get("CharOverlineColor")

    @prop_overline_color.setter
    def prop_overline_color(self, value: Color | None) -> None:
        if value is None:
            self._remove("CharOverlineColor")
            return
        self._set("CharOverlineColor", value)

    @property
    def prop_underine_color(self) -> Color | None:
        """This property specifies if the property ``CharUnderlineColor`` is used for an underline."""
        return self._get("CharUnderlineColor")

    @prop_underine_color.setter
    def prop_underine_color(self, value: Color | None) -> None:
        if value is None:
            self._remove("CharUnderlineColor")
            return
        self._set("CharUnderlineColor", value)

    @property
    def prop_rotation(self) -> float | None:
        """
        This optional property determines the rotation of a character in degrees.

        Depending on the implementation only certain values may be allowed.
        """
        pv = cast(int, self._get("CharRotation"))
        if not pv is None:
            return float(pv / 10)
        return None

    @prop_rotation.setter
    def prop_rotation(self, value: float | None) -> None:
        if value is None:
            self._remove("CharRotation")
            return
        self._set("CharRotation", round(value * 10))

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
