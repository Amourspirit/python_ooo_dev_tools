# region Imports
from __future__ import annotations
from typing import cast
import uno

from ooodev.units import UnitObj
from ooodev.format.writer.style.char.kind.style_char_kind import StyleCharKind as StyleCharKind
from ooodev.format.inner.direct.write.char.font.font_only import FontOnly as InnerFontOnly
from ooodev.format.inner.direct.write.char.font.font_only import FontLang as FontLang
from ..char_style_base_multi import CharStyleBaseMulti

# endregion Imports


class FontOnly(CharStyleBaseMulti):
    """
    Character Style Font

    .. seealso::

        - :ref:`help_writer_format_modify_char_font_only`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        name: str | None = None,
        size: float | UnitObj | None = None,
        font_style_name: str | None = None,
        lang: FontLang | None = None,
        style_name: StyleCharKind | str = StyleCharKind.STANDARD,
        style_family: str = "CharacterStyles",
    ) -> None:
        """
        Constructor

        Args:
            name (str, optional): This property specifies the name of the font style. It may contain more than one name separated by comma.
            size (float, optional): This value contains the size of the characters in ``pt`` (point) units or :ref:`proto_unit_obj`.
            font_style_name (str, optional): Font style name such as ``Bold``.
            lang (Lang, optional): Font Language
            shadowed (bool, optional): Specifies if the characters are formatted and displayed with a shadow effect.
            style_name (StyleParaKind, str, optional): Specifies the Character Style that instance applies to. Default is Default Character Style.
            style_family (str, optional): Style family. Default ``CharacterStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_writer_format_modify_char_font_only`
        """

        direct = InnerFontOnly(name=name, size=size, font_style=font_style_name, lang=lang)
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
    ) -> FontOnly:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Document Object.
            style_name (StyleCharKind, str, optional): Specifies the Character Style that instance applies to. Default is Default Character Style.
            style_family (str, optional): Style family. Default ``CharacterStyles``.

        Returns:
            FontOnly: ``FontOnly`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerFontOnly.from_obj(inst.get_style_props(doc))
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
    def prop_inner(self) -> InnerFontOnly:
        """Gets/Sets Inner Font instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerFontOnly, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerFontOnly) -> None:
        if not isinstance(value, InnerFontOnly):
            raise TypeError(f'Expected type of InnerFontOnly, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())
