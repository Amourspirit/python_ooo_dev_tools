from __future__ import annotations
from typing import cast
import uno
from ....writer.style.char.kind.style_char_kind import StyleCharKind as StyleCharKind
from ..char_style_base_multi import CharStyleBaseMulti
from ....direct.char.border.shadow import Shadow as DirectShadow
from .....utils.color import Color, StandardColor

from ooo.dyn.table.shadow_format import ShadowFormat as ShadowFormat
from ooo.dyn.table.shadow_location import ShadowLocation as ShadowLocation


class Shadow(CharStyleBaseMulti):
    """
    Character Style Border padding.

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        location: ShadowLocation = ShadowLocation.BOTTOM_RIGHT,
        color: Color = StandardColor.GRAY,
        transparent: bool = False,
        width: float = 1.76,
        style_name: StyleCharKind | str = StyleCharKind.STANDARD,
        style_family: str = "CharacterStyles",
    ) -> None:
        """
        Constructor

        Args:
            location (ShadowLocation, optional): contains the location of the shadow. Default to ``ShadowLocation.BOTTOM_RIGHT``.
            color (Color, optional):contains the color value of the shadow. Defaults to ``StandardColor.GRAY``.
            transparent (bool, optional): Shadow transparency. Defaults to False.
            width (float, optional): contains the size of the shadow (in mm units). Defaults to ``1.76``.
            style_name (StyleCharKind, str, optional): Specifies the Character Style that instance applies to. Deftult is Default Character Style.
            style_family (str, optional): Style family. Defatult ``CharacterStyles``.

        Returns:
            None:
        """

        direct = DirectShadow(location=location, color=color, transparent=transparent, width=width)
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
    ) -> Shadow:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Documnet Object.
            style_name (StyleParaKind, str, optional): Specifies the Character Style that instance applies to. Deftult is Default Character Style.
            style_family (str, optional): Style family. Defatult ``CharacterStyles``.

        Returns:
            Shadow: ``Shadow`` instance from document properties.
        """
        inst = super(Shadow, cls).__new__(cls)
        inst.__init__(style_name=style_name, style_family=style_family)
        direct = DirectShadow.from_obj(inst.get_style_props(doc))
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
    def prop_inner(self) -> DirectShadow:
        """Gets Inner Shadow instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(DirectShadow, self._get_style_inst("direct"))
        return self._direct_inner