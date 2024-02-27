# region Import
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno

from ooodev.format.inner.direct.write.char.font.font_position import CharSpacingKind
from ooodev.format.inner.direct.write.char.font.font_position import FontPosition as InnerFontPosition
from ooodev.format.inner.direct.write.char.font.font_position import FontScriptKind
from ooodev.format.inner.modify.write.para.para_style_base_multi import ParaStyleBaseMulti
from ooodev.format.writer.style.para.kind.style_para_kind import StyleParaKind
from ooodev.units.angle import Angle
from ooodev.utils.data_type.intensity import Intensity

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT
# endregion Import


class FontPosition(ParaStyleBaseMulti):
    """
    Character Style Font Effects

    .. seealso::

        - :ref:`help_writer_format_modify_para_font_position`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        script_kind: FontScriptKind | None = None,
        raise_lower: int | None = None,
        rel_size: int | Intensity | None = None,
        rotation: int | Angle | None = None,
        scale: int | None = None,
        fit: bool | None = None,
        spacing: CharSpacingKind | float | UnitT | None = None,
        pair: bool | None = None,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> None:
        """
        Constructor

        Args:
            script_kind (FontScriptKind, optional): Specifies Superscript/Subscript option.
            raise_lower (int, optional): Specifies raise or Lower as percent value. Set to a value of 0 for automatic.
            rel_size (int, Intensity, optional): Specifies relative Font Size as percent value.
                Set this value to ``0`` for automatic; Otherwise value from ``1`` to ``100``.
            rotation (int, Angle, optional): Specifies the rotation of a character in degrees. Depending on the
                implementation only certain values may be allowed.
            scale (int, optional): Specifies scale width as percent value. Min value is ``1``.
            fit (bool, optional): Specifies if rotation is fit to line.
            spacing (CharSpacingKind, float, UnitT, optional): Specifies character spacing in ``pt`` (point) units
                or :ref:`proto_unit_obj`.
            pair (bool, optional): Specifies pair kerning.
            style_name (StyleParaKind, str, optional): Specifies the Character Style that instance applies to.
                Default is Default Character Style.
            style_family (str, optional): Style family. Default ``ParagraphStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_writer_format_modify_para_font_position`
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
        doc: Any,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> FontPosition:
        """
        Gets instance from Document.

        Args:
            doc (Any): UNO Document Object.
            style_name (StyleParaKind, str, optional): Specifies the Character Style that instance applies to.
                Default is Default Character Style.
            style_family (str, optional): Style family. Default ``ParagraphStyles``.

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
    def prop_style_name(self, value: str | StyleParaKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> InnerFontPosition:
        """Gets/Sets Inner Font Position instance"""
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
