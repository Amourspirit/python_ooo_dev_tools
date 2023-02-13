from __future__ import annotations
from typing import cast
import uno

from ....writer.style.para.kind.style_para_kind import StyleParaKind as StyleParaKind
from ..para_style_base_multi import ParaStyleBaseMulti
from .....utils.data_type.intensity import Intensity as Intensity
from .....utils.data_type.angle import Angle as Angle
from ....direct.char.font.font_position import (
    FontPosition as DirectFontPosition,
    CharSpacingKind as CharSpacingKind,
    FontScriptKind as FontScriptKind,
)


class FontPosition(ParaStyleBaseMulti):
    """
    Character Style Font Effects

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        script_kind: FontScriptKind | None = None,
        raise_lower: int | Intensity | None = None,
        rel_size: int | None = None,
        rotation: int | Angle | None = None,
        scale: int | None = None,
        fit: bool | None = None,
        spacing: CharSpacingKind | float | None = None,
        pair: bool | None = None,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> None:
        """
        Constructor

        Args:
            script_kind (FontScriptKind, optional): Specifies Superscript/Subscript option.
            raise_lower (int, Intensity, optional): Specifies raise or Lower.
            rel_size (int, optional): Specifies realitive Font Size. Set this value to ``0`` for automatic.
            rotation (int, Angle, optional): Specifies the rotation of a character in degrees. Depending on the implementation only certain values may be allowed.
            scale (int, optional): Specifies scale width.
            fit (bool, optional): Specifies if rotation is fit to line.
            spacing (float, optional): Specifies character spacing in point units.
            pair (bool, optional): Specifies pair kerning.
            style_name (StyleParaKind, str, optional): Specifies the Character Style that instance applies to. Deftult is Default Character Style.
            style_family (str, optional): Style family. Defatult ``ParagraphStyles``.

        Returns:
            None:
        """

        direct = DirectFontPosition(
            script_kind=script_kind,
            raise_lower=raise_lower,
            rel_size=rel_size,
            rotation=rotation,
            fit=fit,
            spacing=spacing,
            pair=pair,
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
    ) -> FontPosition:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Documnet Object.
            style_name (StyleParaKind, str, optional): Specifies the Character Style that instance applies to. Deftult is Default Character Style.
            style_family (str, optional): Style family. Defatult ``ParagraphStyles``.

        Returns:
            FontPosition: ``FontPosition`` instance from document properties.
        """
        inst = super(FontPosition, cls).__new__(cls)
        inst.__init__(style_name=style_name, style_family=style_family)
        direct = DirectFontPosition.from_obj(inst.get_style_props(doc))
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
    def prop_inner(self) -> DirectFontPosition:
        """Gets Inner Font Postiion instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(DirectFontPosition, self._get_style_inst("direct"))
        return self._direct_inner
