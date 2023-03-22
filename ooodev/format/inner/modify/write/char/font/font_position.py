# region Imports
from __future__ import annotations
from typing import cast
import uno

from ooodev.format.writer.style.char.kind.style_char_kind import StyleCharKind as StyleCharKind
from ..char_style_base_multi import CharStyleBaseMulti
from ooodev.utils.data_type.intensity import Intensity as Intensity
from ooodev.utils.data_type.angle import Angle as Angle
from ooodev.proto.unit_obj import UnitObj
from ooodev.format.inner.direct.write.char.font.font_position import FontPosition as InnerFontPosition
from ooodev.format.inner.direct.write.char.font.font_position import CharSpacingKind as CharSpacingKind
from ooodev.format.inner.direct.write.char.font.font_position import FontScriptKind as FontScriptKind
# endregion Imports

class FontPosition(CharStyleBaseMulti):
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
        spacing: CharSpacingKind | float | UnitObj | None = None,
        pair: bool | None = None,
        style_name: StyleCharKind | str = StyleCharKind.STANDARD,
        style_family: str = "CharacterStyles",
    ) -> None:
        """
        Constructor

        Args:
            script_kind (FontScriptKind, optional): Specifies Superscript/Subscript option.
            raise_lower (int, Intensity, optional): Specifies raise or Lower as percent value. Min value is ``1``.
            rel_size (int, optional): Specifies realitive Font Size as percent value. Set this value to ``0`` for automatic; Otherwise value from ``1`` to ``100``.
            rotation (int, Angle, optional): Specifies the rotation of a character in degrees. Depending on the implementation only certain values may be allowed.
            scale (int, optional): Specifies scale width as percent value. Min value is ``1``.
            fit (bool, optional): Specifies if rotation is fit to line.
            spacing (CharSpacingKind, float, UnitObj, optional): Specifies character spacing in ``pt`` (point) units or :ref:`proto_unit_obj`.
            pair (bool, optional): Specifies pair kerning.
            style_name (StyleCharKind, str, optional): Specifies the Character Style that instance applies to. Deftult is Default Character Style.
            style_family (str, optional): Style family. Defatult ``CharacterStyles``.

        Returns:
            None:
        """

        direct = InnerFontPosition(
            script_kind=script_kind,
            raise_lower=raise_lower,
            rel_size=rel_size,
            rotation=rotation,
            scale=scale,
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
        style_name: StyleCharKind | str = StyleCharKind.STANDARD,
        style_family: str = "CharacterStyles",
    ) -> FontPosition:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Documnet Object.
            style_name (StyleParaKind, str, optional): Specifies the Character Style that instance applies to. Deftult is Default Character Style.
            style_family (str, optional): Style family. Defatult ``CharacterStyles``.

        Returns:
            FontPosition: ``FontPosition`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerFontPosition.from_obj(inst.get_style_props(doc))
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
    def prop_inner(self) -> InnerFontPosition:
        """Gets/Sets Inner Font Postiion instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerFontPosition, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerFontPosition) -> None:
        if not isinstance(value, InnerFontPosition):
            raise TypeError(f'Expected type of InnerFontPosition, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())
