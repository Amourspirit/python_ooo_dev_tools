from __future__ import annotations
from typing import cast
from .sytle_base import StyleBase
from ..utils.color import Color

from ooo.dyn.awt.font_strikeout import FontStrikeoutEnum as FontStrikeoutEnum
from ooo.dyn.awt.font_underline import FontUnderlineEnum as FontUnderlineEnum
from ooo.dyn.awt.font_weight import FontWeightEnum as FontWeightEnum
from ooo.dyn.awt.char_set import CharSetEnum as CharSetEnum
from ooo.dyn.awt.font_family import FontFamilyEnum as FontFamilyEnum


class Font(StyleBase):
    def __init__(
        self,
        name: str | None = None,
        charset: CharSetEnum | None = None,
        family: FontFamilyEnum | None = None,
        b: bool | None = None,
        color: Color | None = None,
        strike: FontStrikeoutEnum | None = None,
        underine: FontUnderlineEnum | None = None,
        underine_color: Color | None = None,
        weight: FontWeightEnum | None = None,
    ) -> None:
        """
        Font options used in styles.

        Args:
            name (str | None, optional): This property specifies the name of the font style. It may contain more than one name separated by comma.
            charset (CharSetEnum | None, optional): The text encoding of the font.
            family (FontFamilyEnum | None, optional): Font Family
            b (bool | None, optional): Short cut to set ``weight`` to bold.
            color (Color | None, optional): The value of the text color.
            strike (FontStrikeoutEnum | None, optional): Dermines the type of the strike out of the character.
            underine (FontUnderlineEnum | None, optional): The value for the character underline.
            underine_color (Color | None, optional): Specifies if the property ``CharUnderlineColor`` is used for an underline.
            weight (FontWeightEnum | None, optional): The value of the font weight.
        """
        # see: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1CharacterProperties.html
        init_vals = {
            "FontName": name,
            "CharColor": color,
            "CharUnderlineColor": underine_color,
        }
        if not charset is None:
            init_vals["CharFontCharSet"] = charset.value
        if not family is None:
            init_vals["CharFontFamily"] = family.value
        if not strike is None:
            init_vals["CharStrikeout"] = strike.value
        if not underine is None:
            init_vals["CharUnderline"] = underine.value
        if not b is None:
            if b:
                init_vals["CharWeight"] = FontWeightEnum.BOLD.value
            else:
                init_vals["CharWeight"] = FontWeightEnum.NORMAL.value
        if not weight is None:
            init_vals["CharWeight"] = weight.value
        super().__init__(**init_vals)

    @property
    def b(self) -> bool | None:
        """Specifies bold"""
        pv = cast(float, self._get("CharWeight"))
        if not pv is None:
            return pv == FontWeightEnum.BOLD.value
        return None

    @property
    def charset(self) -> CharSetEnum | None:
        """This property contains the text encoding of the font."""
        pv = cast(int, self._get("CharFontCharSet"))
        if not pv is None:
            return pv == CharSetEnum(pv)
        return None

    @property
    def color(self) -> Color | None:
        """This property contains the value of the text color."""
        return self._color

    @property
    def family(self) -> FontFamilyEnum | None:
        """This property contains font family."""
        pv = cast(FontFamilyEnum, self._get("CharFontFamily"))
        if not pv is None:
            return FontFamilyEnum(pv)
        return None

    @property
    def name(self) -> str | None:
        """This property specifies the name of the font style. It may contain more than one name separated by comma."""
        return self._get("FontName")

    @property
    def strike(self) -> FontStrikeoutEnum | None:
        """This property determines the type of the strike out of the character."""
        pv = cast(int, self._get("CharStrikeout"))
        if not pv is None:
            FontStrikeoutEnum(pv)
        return None

    @property
    def weight(self) -> FontWeightEnum | None:
        """This property contains the value of the font weight."""
        pv = cast(float, self._get("CharWeight"))
        if not pv is None:
            FontWeightEnum(pv)
        return None

    @property
    def underline(self) -> FontUnderlineEnum | None:
        """This property contains the value for the character underline."""
        pv = cast(int, self._get("CharUnderline"))
        if not pv is None:
            FontUnderlineEnum(pv)
        return None

    @property
    def underine_color(self) -> Color | None:
        """This property specifies if the property ``CharUnderlineColor`` is used for an underline."""
        return self._get("CharUnderlineColor")
