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
from .....direct.common.props.transparent_gradient_props import TransparentGradientProps
from .....direct.fill.transparent.gradient import Gradient as FillGradient


class HeaderTransparentGradient(FillGradient):
    """
    Header Transparent Gradient

    .. versionadded:: 0.9.0
    """

    def _supported_services(self) -> Tuple[str, ...]:
        return ("com.sun.star.style.PageProperties", "com.sun.star.style.PageStyle")

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.DOC | FormatKind.STYLE

    @property
    def _props(self) -> TransparentGradientProps:
        try:
            return self._props_transparent_gradient
        except AttributeError:
            self._props_transparent_gradient = TransparentGradientProps(
                transparence="HeaderFillTransparence",
                name="HeaderFillTransparenceGradientName",
                struct_prop="HeaderFillTransparenceGradient",
            )
        return self._props_transparent_gradient


class Gradient(PageStyleBaseMulti):
    """
    Page Header Transparent Gradient

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        style: GradientStyle = GradientStyle.LINEAR,
        offset: Offset = Offset(50, 50),
        angle: Angle | int = 0,
        border: Intensity | int = 0,
        grad_intensity: IntensityRange = IntensityRange(100, 100),
        style_name: StylePageKind | str = StylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            style (GradientStyle, optional): Specifies the style of the gradient. Defaults to ``GradientStyle.LINEAR``.
            step_count (int, optional): Specifies the number of steps of change color. Defaults to ``0``.
            offset (offset, optional): Specifies the X-coordinate (start) and Y-coordinate (end), where the gradient begins.
                X is effectively the center of the ``RADIAL``, ``ELLIPTICAL``, ``SQUARE`` and ``RECT`` style gradients. Defaults to ``Offset(50, 50)``.
            angle (Angle, int, optional): Specifies angle of the gradient. Defaults to 0.
            border (int, optional): Specifies percent of the total width where just the start color is used. Defaults to 0.
            grad_intensity (IntensityRange, optional): Specifies the intensity at the start point and stop point of the gradient. Defaults to ``IntensityRange(0, 0)``.
            style_name (StyleParaKind, str, optional): Specifies the Page Style that instance applies to. Deftult is Default Page Style.
            style_family (str, optional): Style family. Defatult ``PageStyles``.

        Returns:
            None:
        """

        direct = HeaderTransparentGradient(
            style=style,
            offset=offset,
            angle=angle,
            border=border,
            grad_intensity=grad_intensity,
        )
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
    ) -> Gradient:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Documnet Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``PageStyles``.

        Returns:
            Gradient: ``Gradient`` instance from document properties.
        """
        inst = super(Gradient, cls).__new__(cls)
        inst.__init__(style_name=style_name, style_family=style_family)
        direct = HeaderTransparentGradient.from_obj(inst.get_style_props(doc))
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
    def prop_inner(self) -> HeaderTransparentGradient:
        """Gets Inner Gradient instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(HeaderTransparentGradient, self._get_style_inst("direct"))
        return self._direct_inner
