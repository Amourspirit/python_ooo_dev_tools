from __future__ import annotations
from typing import Tuple, cast
import uno
from ooo.dyn.table.shadow_location import ShadowLocation as ShadowLocation

from .....proto.unit_obj import UnitObj
from .....utils.color import StandardColor, Color
from ....writer.style.page.kind.writer_style_page_kind import WriterStylePageKind as WriterStylePageKind
from ..page_style_base_multi import PageStyleBaseMulti
from ....direct.structs.shadow_struct import ShadowStruct as InnerShadow

# from ....direct.para.border.shadow import Shadow as DirectShadow


class Shadow(PageStyleBaseMulti):
    """
    Page Style Border Shadow

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        location: ShadowLocation = ShadowLocation.BOTTOM_RIGHT,
        color: Color = StandardColor.GRAY,
        transparent: bool = False,
        width: float | UnitObj = 1.76,
        style_name: WriterStylePageKind | str = WriterStylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            location (ShadowLocation, optional): contains the location of the shadow. Default to ``ShadowLocation.BOTTOM_RIGHT``.
            color (Color, optional):contains the color value of the shadow. Defaults to ``StandardColor.GRAY``.
            transparent (bool, optional): Shadow transparency. Defaults to False.
            width (float, UnitObj, optional): contains the size of the shadow (in ``mm`` units) or :ref:`proto_unit_obj`. Defaults to ``1.76``.
            style_name (StyleParaKind, str, optional): Specifies the Page Style that instance applies to. Deftult is Default Page Style.
            style_family (str, optional): Style family. Defatult ``PageStyles``.

        Returns:
            None:
        """
        cattribs = {"_property_name": "ShadowFormat", "_supported_services_values": self._supported_services()}
        direct = InnerShadow(location=location, color=color, transparent=transparent, width=width, _cattribs=cattribs)
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())

    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: WriterStylePageKind | str = WriterStylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> Shadow:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Documnet Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``PageStyles``.

        Returns:
            Shadow: ``Shadow`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        cattribs = {"_property_name": "ShadowFormat", "_supported_services_values": inst._supported_services()}
        direct = InnerShadow.from_obj(inst.get_style_props(doc), _cattribs=cattribs)
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | WriterStylePageKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> InnerShadow:
        """Gets/Sets Inner Shadow instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerShadow, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerShadow) -> None:
        if not isinstance(value, InnerShadow):
            raise TypeError(f'Expected type of InnerShadow, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())
