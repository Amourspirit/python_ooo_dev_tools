from __future__ import annotations
from typing import cast
from .sytle_base import StyleBase
from ..utils.color import Color

from ooo.dyn.awt.font_strikeout import FontStrikeoutEnum as FontStrikeoutEnum
from ooo.dyn.awt.font_underline import FontUnderlineEnum as FontUnderlineEnum
from ooo.dyn.awt.font_weight import FontWeightEnum as FontWeightEnum
from ooo.dyn.awt.char_set import CharSetEnum as CharSetEnum


class Font(StyleBase):
    def __init__(
        self,
        name: str | None = None,
        charset: CharSetEnum | None = None,
        family: str | None = None,
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
            charset (str | None, optional): This property contains the text encoding of the font.
            family (str | None, optional): _description_. Defaults to None.
            b (bool | None, optional): Short cut to set ``weight`` to bold.
            color (Color | None, optional): _description_. Defaults to None.
            strike (FontStrikeoutEnum | None, optional): _description_. Defaults to None.
            underine (FontUnderlineEnum | None, optional): _description_. Defaults to None.
            underine_color (Color | None, optional): _description_. Defaults to None.
            weight (FontWeightEnum | None, optional): _description_. Defaults to None.
        """
        # see: https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1CharacterProperties.html
        init_vals = {
            "FontName": name,
            "FontFamily": family,
            "CharColor": color,
            "CharUnderlineColor": underine_color,
        }
        if not charset is None:
            init_vals["CharFontCharSet"] = charset.value
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
        """Gets color value"""
        return self._color

    @property
    def family(self) -> str | None:
        """Specifies family"""
        return self._get("FontFamily")

    @property
    def name(self) -> str | None:
        """This property specifies the name of the font style. It may contain more than one name separated by comma."""
        return self._get("FontName")

    @property
    def strike(self) -> FontStrikeoutEnum | None:
        """Gets strike value"""
        pv = cast(int, self._get("CharStrikeout"))
        if not pv is None:
            FontStrikeoutEnum(pv)
        return None

    @property
    def weight(self) -> FontWeightEnum | None:
        """This property contains the value of the font weight."""
        pv = cast(int, self._get("CharWeight"))
        if not pv is None:
            FontWeightEnum(pv)
        return None

    @property
    def underine_color(self) -> Color | None:
        """This property specifies if the property ``CharUnderlineColor`` is used for an underline."""
        return self._get("CharUnderlineColor")
