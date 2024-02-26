"""
Font Only Style for Draw Documents

.. versionadded:: 0.17.9
"""

# region Imports
from __future__ import annotations
from typing import cast, Any, TYPE_CHECKING
import uno

from ooodev.format.inner.direct.write.char.font.font_only import FontOnly as InnerFontOnly
from ooodev.format.inner.direct.write.char.font.font_only import FontLang as FontLang
from ooodev.format.inner.modify.draw.char_style_base_multi import CharStyleBaseMulti
from ooodev.format.draw.style.kind import DrawStyleFamilyKind
from ooodev.format.draw.style.lookup import FamilyGraphics

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT

# endregion Imports


class FontOnly(CharStyleBaseMulti):
    """
    Character Style Font

    .. seealso::

        - :ref:`help_draw_format_modify_font_font_only`

    .. versionadded:: 0.17.9
    """

    def __init__(
        self,
        *,
        name: str | None = None,
        size: float | UnitT | None = None,
        font_style_name: str | None = None,
        lang: FontLang | None = None,
        style_name: str = FamilyGraphics.DEFAULT_DRAWING_STYLE,
        style_family: str | DrawStyleFamilyKind = DrawStyleFamilyKind.GRAPHICS,
    ) -> None:
        """
        Constructor

        Args:
            name (str, optional): This property specifies the name of the font style. It may contain more than one name separated by comma.
            size (float, UnitT, optional): This value contains the size of the characters in ``pt`` (point) units or :ref:`proto_unit_obj`.
            font_style_name (str, optional): Font style name such as ``Bold``.
            lang (Lang, optional): Font Language
            shadowed (bool, optional): Specifies if the characters are formatted and displayed with a shadow effect.
            style_name (FamilyGraphics, str, optional): Specifies the Style that instance applies to.
                Default is Default ``standard`` Style.
            style_family (str, DrawStyleFamilyKind, optional): Family Style. Defaults to ``graphics``.

        Returns:
            None:

        See Also:
            - :ref:`help_draw_format_modify_font_font_only`
        """

        direct = InnerFontOnly(name=name, size=size, font_style=font_style_name, lang=lang)
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = str(style_family)
        self._set_style("direct", direct)

    @classmethod
    def from_style(
        cls,
        doc: Any,
        style_name: str = FamilyGraphics.DEFAULT_DRAWING_STYLE,
        style_family: str | DrawStyleFamilyKind = DrawStyleFamilyKind.GRAPHICS,
    ) -> FontOnly:
        """
        Gets instance from Document.

        Args:
            doc (Any): UNO Document Object.
            style_name (FamilyGraphics, str, optional): Specifies the Style that instance applies to.
                Default is ``FamilyGraphics.DEFAULT_DRAWING_STYLE``.
            style_family (DrawStyleFamilyKind, str, optional): Style family. Default ``DrawStyleFamilyKind.GRAPHICS``.

        Returns:
            FontOnly: ``FontOnly`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerFontOnly.from_obj(inst.get_style_props(doc))
        inst._set_style("direct", direct)
        return inst

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | FamilyGraphics):
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
