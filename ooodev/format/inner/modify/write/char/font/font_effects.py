# region Imports
from __future__ import annotations
from typing import cast
import uno

from ooodev.format.writer.style.char.kind.style_char_kind import StyleCharKind as StyleCharKind
from ooodev.format.inner.modify.write.char.char_style_base_multi import CharStyleBaseMulti
from ooodev.format.inner.direct.write.char.font.font_effects import FontEffects as InnerFontEffects
from ooodev.format.inner.direct.write.char.font.font_effects import FontLine as FontLine
from ooodev.utils.data_type.intensity import Intensity as Intensity
from ooodev.utils.color import Color

from ooo.dyn.awt.font_strikeout import FontStrikeoutEnum as FontStrikeoutEnum
from ooo.dyn.awt.font_underline import FontUnderlineEnum as FontUnderlineEnum
from ooo.dyn.style.case_map import CaseMapEnum as CaseMapEnum
from ooo.dyn.awt.font_relief import FontReliefEnum as FontReliefEnum

# endregion Imports


class FontEffects(CharStyleBaseMulti):
    """
    Character Style Font Effects

    .. seealso::

        - :ref:`help_writer_format_modify_char_font_effects`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        color: Color | None = None,
        transparency: Intensity | int | None = None,
        overline: FontLine | None = None,
        underline: FontLine | None = None,
        strike: FontStrikeoutEnum | None = None,
        word_mode: bool | None = None,
        case: CaseMapEnum | None = None,
        relief: FontReliefEnum | None = None,
        outline: bool | None = None,
        hidden: bool | None = None,
        shadowed: bool | None = None,
        style_name: StyleCharKind | str = StyleCharKind.STANDARD,
        style_family: str = "CharacterStyles",
    ) -> None:
        """
        Constructor

        Args:
            color (:py:data:`~.utils.color.Color`, optional): The value of the text color.
                If value is ``-1`` the automatic color is applied.
            transparency (Intensity, int, optional): The transparency value from ``0`` to ``100`` for the font color.
            overline (FontLine, optional): Character overline values.
            underline (FontLine, optional): Character underline values.
            strike (FontStrikeoutEnum, optional): Determines the type of the strike out of the character.
            word_mode(bool, optional): If ``True``, the underline and strike-through properties are not applied to white spaces.
            case (CaseMapEnum, optional): Specifies the case of the font.
            relief (FontReliefEnum, optional): Specifies the relief of the font.
            outline (bool, optional): Specifies if the font is outlined.
            hidden (bool, optional): Specifies if the font is hidden.
            shadowed (bool, optional): Specifies if the characters are formatted and displayed with a shadow effect.
            style_name (StyleParaKind, str, optional): Specifies the Character Style that instance applies to. Default is Default Character Style.
            style_family (str, optional): Style family. Default ``CharacterStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_writer_format_modify_char_font_effects`
        """

        direct = InnerFontEffects(
            color=color,
            transparency=transparency,
            overline=overline,
            underline=underline,
            strike=strike,
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
        style_name: StyleCharKind | str = StyleCharKind.STANDARD,
        style_family: str = "CharacterStyles",
    ) -> FontEffects:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Document Object.
            style_name (StyleCharKind, str, optional): Specifies the Character Style that instance applies to. Default is Default Character Style.
            style_family (str, optional): Style family. Default ``CharacterStyles``.

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
    def prop_style_name(self, value: str | StyleCharKind):
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
