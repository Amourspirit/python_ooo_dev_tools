from __future__ import annotations
from typing import cast
from .sytle_base import StyleBase
from ..utils.color import Color
from ..utils import info as mInfo
from ..utils import lo as mLo

from ooo.dyn.awt.font_strikeout import FontStrikeoutEnum as StrikeOutKind
from ooo.dyn.awt.font_underline import FontUnderlineEnum as UnderlineKind
from ooo.dyn.awt.font_weight import FontWeightEnum as WeightKind
from ooo.dyn.awt.char_set import CharSetEnum as CharSetKnid
from ooo.dyn.awt.font_family import FontFamilyEnum as FamilyKind
from ooo.dyn.awt.font_slant import FontSlant as SlantKind


class Font(StyleBase):
    def __init__(
        self,
        name: str | None = None,
        charset: CharSetKnid | None = None,
        family: FamilyKind | None = None,
        b: bool | None = None,
        i: bool | None = None,
        u: bool | None = None,
        color: Color | None = None,
        strike: StrikeOutKind | None = None,
        underine: UnderlineKind | None = None,
        underine_color: Color | None = None,
        weight: WeightKind | None = None,
        slant: SlantKind | None = None,
        super_script: bool | None = None,
        sub_script: bool | None = None,
    ) -> None:
        """
        Font options used in styles.

        Args:
            name (str | None, optional): This property specifies the name of the font style. It may contain more than one name separated by comma.
            charset (CharSetKnid | None, optional): The text encoding of the font.
            family (FamilyKind | None, optional): Font Family
            b (bool | None, optional): Short cut to set ``weight`` to bold.
            i (bool | None, optional): Short cut to set ``slant`` to italic.
            u (bool | None, optional): Short cut ot set ``underline`` to underline.
            color (Color | None, optional): The value of the text color.
            strike (StrikeOutKind | None, optional): Dermines the type of the strike out of the character.
            underine (UnderlineKind | None, optional): The value for the character underline.
            underine_color (Color | None, optional): Specifies if the property ``CharUnderlineColor`` is used for an underline.
            weight (WeightKind | None, optional): The value of the font weight.
            slant (SlantKind | None, optional): The value of the posture of the document such as ``SlantKind.ITALIC``.
            super_script (bool | None, optional): Super script option.
            sub_script (bool | None, optional): Sub script option.
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

        if not b is None:
            if b:
                init_vals["CharWeight"] = WeightKind.BOLD.value
            else:
                init_vals["CharWeight"] = WeightKind.NORMAL.value
        if not i is None:
            if i:
                init_vals["CharPosture"] = SlantKind.ITALIC
            else:
                init_vals["CharWeight"] = SlantKind.NONE
        if not u is None:
            if u:
                init_vals["CharUnderline"] = UnderlineKind.SINGLE.value
            else:
                init_vals["CharUnderline"] = UnderlineKind.NONE.value

        if not underine is None:
            init_vals["CharUnderline"] = underine.value

        if not weight is None:
            init_vals["CharWeight"] = weight.value

        if not slant is None:
            init_vals["CharPosture"] = slant

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

        super().__init__(**init_vals)

    def apply_style(self, obj: object) -> None:
        if mInfo.Info.support_service(obj, "com.sun.star.style.CharacterProperties"):
            super().apply_style(obj)
        else:
            mLo.Lo.print("Unable to apply font style. CharacterProperties service not supported")

    @property
    def b(self) -> bool:
        """Specifies bold"""
        pv = cast(float, self._get("CharWeight"))
        if not pv is None:
            return pv == WeightKind.BOLD.value
        return False

    @property
    def i(self) -> bool:
        """Specifies italic"""
        pv = cast(SlantKind, self._get("CharPosture"))
        if not pv is None:
            return pv == SlantKind.ITALIC
        return False

    @property
    def u(self) -> bool:
        """Specifies underline"""
        pv = cast(int, self._get("CharUnderline"))
        if not pv is None:
            return pv != UnderlineKind.NONE.value
        return False

    @property
    def charset(self) -> CharSetKnid | None:
        """This property contains the text encoding of the font."""
        pv = cast(int, self._get("CharFontCharSet"))
        if not pv is None:
            return CharSetKnid(pv)
        return None

    @property
    def color(self) -> Color | None:
        """This property contains the value of the text color."""
        return self._get("CharColor")

    @property
    def family(self) -> FamilyKind | None:
        """This property contains font family."""
        pv = cast(FamilyKind, self._get("CharFontFamily"))
        if not pv is None:
            return FamilyKind(pv)
        return None

    @property
    def name(self) -> str | None:
        """This property specifies the name of the font style. It may contain more than one name separated by comma."""
        return self._get("FontName")

    @property
    def strike(self) -> StrikeOutKind | None:
        """This property determines the type of the strike out of the character."""
        pv = cast(int, self._get("CharStrikeout"))
        if not pv is None:
            return StrikeOutKind(pv)
        return None

    @property
    def weight(self) -> WeightKind | None:
        """This property contains the value of the font weight."""
        pv = cast(float, self._get("CharWeight"))
        if not pv is None:
            return WeightKind(pv)
        return None

    @property
    def slant(self) -> SlantKind | None:
        """This property contains the value of the posture of the document such as  ``SlantKind.ITALIC``"""
        return self._get("CharPosture")

    @property
    def super_script(self) -> bool:
        pv = cast(int, self._get("CharEscapement"))
        if not pv is None:
            return pv > 0
        return False

    @property
    def sub_script(self) -> bool:
        pv = cast(int, self._get("CharEscapement"))
        if not pv is None:
            return pv < 0
        return False

    @property
    def underline(self) -> UnderlineKind | None:
        """This property contains the value for the character underline."""
        pv = cast(int, self._get("CharUnderline"))
        if not pv is None:
            return UnderlineKind(pv)
        return None

    @property
    def underine_color(self) -> Color | None:
        """This property specifies if the property ``CharUnderlineColor`` is used for an underline."""
        return self._get("CharUnderlineColor")
