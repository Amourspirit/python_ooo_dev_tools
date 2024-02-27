# region Import
from __future__ import annotations
from typing import cast, Tuple
import uno
from ooo.dyn.awt.font_strikeout import FontStrikeoutEnum
from ooo.dyn.style.case_map import CaseMapEnum
from ooo.dyn.awt.font_relief import FontReliefEnum

from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.calc.style.cell.kind.style_cell_kind import StyleCellKind
from ooodev.format.inner.direct.write.char.font import font_effects
from ooodev.format.inner.direct.write.char.font.font_effects import FontLine
from ooodev.utils.data_type.intensity import Intensity
from ooodev.utils.color import Color
from ooodev.format.inner.modify.calc.cell_style_base_multi import CellStyleBaseMulti

# endregion Import


class InnerFontEffects(font_effects.FontEffects):
    """
    Inner Style Font Effects

    .. versionadded:: 0.9.0
    """

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.style.CellStyle",)
        return self._supported_services_values

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.STYLE
        return self._format_kind_prop


class FontEffects(CellStyleBaseMulti):
    """
    Style Font Effects

    .. seealso::

        - :ref:`help_calc_format_modify_cell_font_effects`

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
        style_name: StyleCellKind | str = StyleCellKind.DEFAULT,
        style_family: str = "CellStyles",
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
            style_name (StyleCellKind, str, optional): Specifies the Cell Style that instance applies to.
                Default is Default Cell Style.
            style_family (str, optional): Style family. Default ``CellStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_modify_cell_font_effects`
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
        style_name: StyleCellKind | str = StyleCellKind.DEFAULT,
        style_family: str = "CellStyles",
    ) -> FontEffects:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Document Object.
            style_name (StyleCellKind, str, optional): Specifies the Cell Style that instance applies to.
                Default is Default Cell Style.
            style_family (str, optional): Style family. Default ``CellStyles``.

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
    def prop_style_name(self, value: str | StyleCellKind):
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
