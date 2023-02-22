from __future__ import annotations
from typing import cast
import uno
from ....writer.style.para.kind.style_para_kind import StyleParaKind as StyleParaKind
from ..para_style_base_multi import ParaStyleBaseMulti
from ....direct.char.font.font_effects import FontEffects as InnerFontEffects
from .....utils.data_type.intensity import Intensity as Intensity
from .....utils.color import Color

from ooo.dyn.awt.font_strikeout import FontStrikeoutEnum as FontStrikeoutEnum
from ooo.dyn.awt.font_underline import FontUnderlineEnum as FontUnderlineEnum
from ooo.dyn.style.case_map import CaseMapEnum as CaseMapEnum
from ooo.dyn.awt.font_relief import FontReliefEnum as FontReliefEnum


class FontEffects(ParaStyleBaseMulti):
    """
    Style Font Effects

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        color: Color | None = None,
        transparency: Intensity | int | None = None,
        overline: FontUnderlineEnum | None = None,
        overline_color: Color | None = None,
        strike: FontStrikeoutEnum | None = None,
        underine: FontUnderlineEnum | None = None,
        underline_color: Color | None = None,
        word_mode: bool | None = None,
        case: CaseMapEnum | None = None,
        relief: FontReliefEnum | None = None,
        outline: bool | None = None,
        hidden: bool | None = None,
        shadowed: bool | None = None,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> None:
        """
        Constructor

        Args:
            color (Color, optional): The value of the text color.
                If value is ``-1`` the automatic color is applied.
            transparency (Intensity, int, optional): The transparency value from ``0`` to ``100`` for the font color.
            overline (FontUnderlineEnum, optional): The value for the character overline.
            overline_color (Color, optional): Specifies if the property ``CharOverlinelineColor`` is used for an overline.
                If value is ``-1`` the automatic color is applied.
            strike (FontStrikeoutEnum, optional): Detrmines the type of the strike out of the character.
            underine (FontUnderlineEnum, optional): The value for the character underline.
            underline_color (Color, optional): Specifies if the property ``CharUnderlineColor`` is used for an underline.
                If value is ``-1`` the automatic color is applied.
            word_mode(bool, optional): If ``True``, the underline and strike-through properties are not applied to white spaces.
            case (CaseMapEnum, optional): Specifies the case of the font.
            releif (FontReliefEnum, optional): Specifies the relief of the font.
            outline (bool, optional): Specifies if the font is outlined.
            hidden (bool, optional): Specifies if the font is hidden.
            shadowed (bool, optional): Specifies if the characters are formatted and displayed with a shadow effect.
            style_name (StyleParaKind, str, optional): Specifies the Character Style that instance applies to. Deftult is Default Character Style.
            style_family (str, optional): Style family. Defatult ``ParagraphStyles``.

        Returns:
            None:
        """

        direct = InnerFontEffects(
            color=color,
            transparency=transparency,
            overline=overline,
            overline_color=overline_color,
            strike=strike,
            underine=underine,
            underline_color=underline_color,
            word_mode=word_mode,
            case=case,
            relief=relief,
            outline=outline,
            hidden=hidden,
            shadowed=shadowed,
        )
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())

    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> FontEffects:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Documnet Object.
            style_name (StyleParaKind, str, optional): Specifies the Character Style that instance applies to. Deftult is Default Character Style.
            style_family (str, optional): Style family. Defatult ``ParagraphStyles``.

        Returns:
            FontEffects: ``FontEffects`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerFontEffects.from_obj(inst.get_style_props(doc))
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | StyleParaKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> InnerFontEffects:
        """Gets/Sets Inner Font Effects instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerFontEffects, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerFontEffects) -> None:
        if not isinstance(value, InnerFontEffects):
            raise TypeError(f'Expected type of InnerFontEffects, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())
