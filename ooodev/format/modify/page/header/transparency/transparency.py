from __future__ import annotations
from typing import Tuple, cast
import uno
from ooo.dyn.awt.gradient_style import GradientStyle as GradientStyle
from ...page_style_base_multi import PageStyleBaseMulti
from ......utils.data_type.angle import Angle as Angle
from ......utils.data_type.color_range import ColorRange as ColorRange
from ......utils.data_type.intensity import Intensity as Intensity
from ......utils.data_type.intensity_range import IntensityRange as IntensityRange
from ......utils.data_type.offset import Offset as Offset
from .....writer.style.page.kind.style_page_kind import StylePageKind as StylePageKind
from .....kind.format_kind import FormatKind
from .....preset.preset_gradient import PresetGradientKind as PresetGradientKind
from .....direct.common.props.transparent_transparency_props import TransparentTransparencyProps
from .....direct.fill.transparent.transparency import Transparency as FillTransparency


class HeaderTransparentTransparency(FillTransparency):
    """
    Header Transparent Transparency

    .. versionadded:: 0.9.0
    """

    def _supported_services(self) -> Tuple[str, ...]:
        return ("com.sun.star.style.PageProperties", "com.sun.star.style.PageStyle")

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.DOC | FormatKind.STYLE

    @property
    def _props(self) -> TransparentTransparencyProps:
        try:
            return self._props_transparency
        except AttributeError:
            self._props_transparency = TransparentTransparencyProps(transparence="HeaderFillTransparence")
        return self._props_transparency


class Transparency(PageStyleBaseMulti):
    """
    Header Transparent Transparency

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        value: Intensity | int = 0,
        style_name: StylePageKind | str = StylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            value (Intensity, int, optional): Specifies the transparency value from ``0`` to ``100``.
            style_name (StyleParaKind, str, optional): Specifies the Page Style that instance applies to. Deftult is Default Page Style.
            style_family (str, optional): Style family. Defatult ``PageStyles``.

        Returns:
            None:
        """

        direct = HeaderTransparentTransparency(value=value)
        direct._prop_parent = self
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())

    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: StylePageKind | str = StylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> Transparency:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Documnet Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``PageStyles``.

        Returns:
            Transparency: ``Transparency`` instance from document properties.
        """
        inst = super(Transparency, cls).__new__(cls)
        inst.__init__(style_name=style_name, style_family=style_family)
        direct = HeaderTransparentTransparency.from_obj(inst.get_style_props(doc))
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | StylePageKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> HeaderTransparentTransparency:
        """Gets Inner Transparency instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(HeaderTransparentTransparency, self._get_style_inst("direct"))
        return self._direct_inner
