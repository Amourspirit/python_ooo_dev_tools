from __future__ import annotations
from typing import cast
from enum import Enum

from ..exceptions import ex as mEx
from ..utils import info as mInfo
from ..utils import lo as mLo
from ..utils.color import Color
from .style_base import StyleBase
from .style_const import POINT_RATIO

from ooo.dyn.awt.char_set import CharSetEnum as CharSetEnum
from ooo.dyn.awt.font_family import FontFamilyEnum as FontFamilyEnum
from ooo.dyn.awt.font_slant import FontSlant as FontSlant
from ooo.dyn.awt.font_strikeout import FontStrikeoutEnum as FontStrikeoutEnum
from ooo.dyn.awt.font_underline import FontUnderlineEnum as FontUnderlineEnum
from ooo.dyn.awt.font_weight import FontWeightEnum as FontWeightEnum
from ooo.dyn.table.shadow_format import ShadowFormat as ShadowFormat


class CharSpacingKind(float, Enum):
    """Character Spacing"""

    VERY_TIGHT = -3.0
    TIGHT = -1.5
    NORMAL = 0.0
    LOOSE = 3.0
    VERY_LOOSE = 6.0


class Font(StyleBase):
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
        sub_script: bool | None = None,
        super_script: bool | None = None,
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
            spacing (CharSpacingKind, float, optional): Character spacing in point units.
            shadowed (bool, optional): Specifies if the characters are formatted and displayed with a shadow effect.
            shadow_fmt: (ShadowFormat, optional): Determines the type, color, and width of the shadow.
            strike (FontStrikeoutEnum, optional): Dermines the type of the strike out of the character.
            sub_script (bool, optional): Sub script option.
            super_script (bool, optional): Super script option.
            underine (FontUnderlineEnum, optional): The value for the character underline.
            underine_color (Color, optional): Specifies if the property ``CharUnderlineColor`` is used for an underline.
            weight (FontWeightEnum, optional): The value of the font weight.
            word_mode(bool, optional): If ``True``, the underline and strike-through properties are not applied to white spaces.
        """
        # could not find any documention in the API or elsewhere online for Overline
        # see: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1CharacterProperties.html
        init_vals = {
            "FontName": name,
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

        if not super_script is None:
            if super_script:
                init_vals["CharEscapementHeight"] = 58
                init_vals["CharEscapement"] = 14_000
            else:
                init_vals["CharEscapementHeight"] = 100
                init_vals["CharEscapement"] = 0

        if not sub_script is None:
            if sub_script:
                init_vals["CharEscapementHeight"] = 58
                init_vals["CharEscapement"] = -14_000
            else:
                init_vals["CharEscapementHeight"] = 100
                init_vals["CharEscapement"] = 0

        if not rotation is None:
            init_vals["CharRotation"] = round(rotation * 10)

        if not shadow_fmt is None:
            if mInfo.Info.is_type_struct(shadow_fmt, "com.sun.star.table.ShadowFormat"):
                init_vals["CharShadowFormat"] = shadow_fmt

        super().__init__(**init_vals)

    def apply_style(self, obj: object) -> None:
        """
        Applies styles to object

        Args:
            obj (object): UNO Oject that styles are to be applied.

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

    @property
    def b(self) -> bool:
        """Specifies bold"""
        pv = cast(float, self._get("CharWeight"))
        if not pv is None:
            return pv == FontWeightEnum.BOLD.value
        return False

    @property
    def bg_color(self) -> Color | None:
        """This property contains the text background color."""
        return self._get("CharBackColor")

    @property
    def bg_color_transparent(self) -> bool | None:
        """This property contains the text background color."""
        return self._get("CharBackTransparent")

    @property
    def i(self) -> bool | None:
        """Specifies italic"""
        pv = cast(FontSlant, self._get("CharPosture"))
        if not pv is None:
            return pv == FontSlant.ITALIC
        return None

    @property
    def u(self) -> bool | None:
        """Specifies underline"""
        pv = cast(int, self._get("CharUnderline"))
        if not pv is None:
            return pv != FontUnderlineEnum.NONE.value
        return None

    @property
    def charset(self) -> CharSetEnum | None:
        """This property contains the text encoding of the font."""
        pv = cast(int, self._get("CharFontCharSet"))
        if not pv is None:
            return CharSetEnum(pv)
        return None

    @property
    def color(self) -> Color | None:
        """This property contains the value of the text color."""
        return self._get("CharColor")

    @property
    def family(self) -> FontFamilyEnum | None:
        """This property contains font family."""
        pv = cast(FontFamilyEnum, self._get("CharFontFamily"))
        if not pv is None:
            return FontFamilyEnum(pv)
        return None

    @property
    def size(self) -> float | None:
        """This value contains the size of the characters in point."""
        return self._get("CharHeight")

    @property
    def name(self) -> str | None:
        """This property specifies the name of the font style. It may contain more than one name separated by comma."""
        return self._get("FontName")

    @property
    def strike(self) -> FontStrikeoutEnum | None:
        """This property determines the type of the strike out of the character."""
        pv = cast(int, self._get("CharStrikeout"))
        if not pv is None:
            return FontStrikeoutEnum(pv)
        return None

    @property
    def weight(self) -> FontWeightEnum | None:
        """This property contains the value of the font weight."""
        pv = cast(float, self._get("CharWeight"))
        if not pv is None:
            return FontWeightEnum(pv)
        return None

    @property
    def slant(self) -> FontSlant | None:
        """This property contains the value of the posture of the document such as  ``FontSlant.ITALIC``"""
        return self._get("CharPosture")

    @property
    def spacing(self) -> float | None:
        """This value contains character spacing in point units"""
        pv = self._get("CharKerning")
        if not pv is None:
            if pv == 0.0:
                return 0.0
            return pv / POINT_RATIO
        return None

    @property
    def shadowed(self) -> bool | None:
        """This property specifies if the characters are formatted and displayed with a shadow effect."""
        pv = cast(int, self._get("CharShadowed"))
        if not pv is None:
            return pv > 0
        return None

    @property
    def shadow_fmt(self) -> ShadowFormat | None:
        """This property specifies the type, color, and width of the shadow."""
        return self._get("CharShadowFormat")

    @property
    def super_script(self) -> bool | None:
        """Specifies if the font is super script"""
        pv = cast(int, self._get("CharEscapement"))
        if not pv is None:
            return pv > 0
        return None

    @property
    def sub_script(self) -> bool | None:
        """Specifies if the font is sub script"""
        pv = cast(int, self._get("CharEscapement"))
        if not pv is None:
            return pv < 0
        return None

    @property
    def overline(self) -> FontUnderlineEnum | None:
        """This property contains the value for the character overline."""
        pv = cast(int, self._get("CharOverline"))
        if not pv is None:
            return FontUnderlineEnum(pv)
        return None

    @property
    def overline_color(self) -> Color | None:
        """This property specifies if the property ``CharOverlineColor`` is used for an underline."""
        return self._get("CharOverlineColor")

    @property
    def underline(self) -> FontUnderlineEnum | None:
        """This property contains the value for the character underline."""
        pv = cast(int, self._get("CharUnderline"))
        if not pv is None:
            return FontUnderlineEnum(pv)
        return None

    @property
    def underine_color(self) -> Color | None:
        """This property specifies if the property ``CharUnderlineColor`` is used for an underline."""
        return self._get("CharUnderlineColor")

    @property
    def rotation(self) -> float | None:
        """
        This optional property determines the rotation of a character in degrees.

        Depending on the implementation only certain values may be allowed.
        """
        pv = cast(int, self._get("CharRotation"))
        if not pv is None:
            return float(pv / 10)
        return None

    @property
    def word_mode(self) -> bool | None:
        """If this property is ``True``, the underline and strike-through properties are not applied to white spaces."""
        return self._get("CharWordMode")
