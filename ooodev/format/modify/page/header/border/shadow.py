from __future__ import annotations
from typing import cast, Type, TypeVar
import uno
from ooo.dyn.table.shadow_location import ShadowLocation as ShadowLocation

from ......utils.color import StandardColor, Color
from .....writer.style.page.kind.style_page_kind import StylePageKind as StylePageKind
from ...page_style_base_multi import PageStyleBaseMulti
from .....direct.structs.shadow_struct import ShadowStruct

_TShadow = TypeVar(name="_TShadow", bound="Shadow")


class Shadow(PageStyleBaseMulti):
    """
    Page Style Header Border Shadow

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        location: ShadowLocation = ShadowLocation.BOTTOM_RIGHT,
        color: Color = StandardColor.GRAY,
        transparent: bool = False,
        width: float = 1.76,
        style_name: StylePageKind | str = StylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            location (ShadowLocation, optional): contains the location of the shadow. Default to ``ShadowLocation.BOTTOM_RIGHT``.
            color (Color, optional):contains the color value of the shadow. Defaults to ``StandardColor.GRAY``.
            transparent (bool, optional): Shadow transparency. Defaults to False.
            width (float, optional): contains the size of the shadow (in mm units). Defaults to ``1.76``.
            style_name (StyleParaKind, str, optional): Specifies the Page Style that instance applies to. Deftult is Default Page Style.
            style_family (str, optional): Style family. Defatult ``PageStyles``.

        Returns:
            None:
        """

        direct = ShadowStruct(
            location=location, color=color, transparent=transparent, width=width, _cattribs=self._get_inner_cattribs()
        )
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())

    # region Internal Methods
    def _get_inner_prop_name(self) -> str:
        return "HeaderShadowFormat"

    def _get_inner_cattribs(self) -> dict:
        return {
            "_supported_services_values": self._supported_services(),
            "_format_kind_prop": self.prop_format_kind,
            "_property_name": self._get_inner_prop_name(),
        }

    # endregion Internal Methods

    # region Static Methods
    @classmethod
    def from_style(
        cls: Type[_TShadow],
        doc: object,
        style_name: StylePageKind | str = StylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> _TShadow:
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
        direct = ShadowStruct.from_obj(inst.get_style_props(doc), _cattribs=inst._get_inner_cattribs())
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    # endregion Static Methods

    # region Properties
    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | StylePageKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> ShadowStruct:
        """Gets Inner Shadow instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(ShadowStruct, self._get_style_inst("direct"))
        return self._direct_inner

    # endregion Properties