# region Import
from __future__ import annotations
from typing import Any, cast, Type, TypeVar
import uno
from ooo.dyn.awt.gradient_style import GradientStyle
from ooodev.format.inner.common.props.area_gradient_props import AreaGradientProps
from ooodev.format.inner.direct.write.fill.area.gradient import Gradient as InnerGradient
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.modify.write.page.page_style_base_multi import PageStyleBaseMulti
from ooodev.format.inner.preset.preset_gradient import PresetGradientKind
from ooodev.format.writer.style.page.kind.writer_style_page_kind import WriterStylePageKind
from ooodev.units.angle import Angle
from ooodev.utils.color import StandardColor
from ooodev.utils.data_type.color_range import ColorRange
from ooodev.utils.data_type.intensity import Intensity
from ooodev.utils.data_type.intensity_range import IntensityRange
from ooodev.utils.data_type.offset import Offset

# endregion Import
_TGradient = TypeVar("_TGradient", bound="Gradient")


class Gradient(PageStyleBaseMulti):
    """
    Page Header Gradient Color

    .. seealso::

        - :ref:`help_writer_format_modify_page_header_area`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        style: GradientStyle = GradientStyle.LINEAR,
        step_count: int = 0,
        offset: Offset = Offset(50, 50),
        angle: Angle | int = 0,
        border: Intensity | int = 0,
        grad_color: ColorRange = ColorRange(StandardColor.BLACK, StandardColor.WHITE),
        grad_intensity: IntensityRange = IntensityRange(100, 100),
        name: str = "",
        style_name: WriterStylePageKind | str = WriterStylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            style (GradientStyle, optional): Specifies the style of the gradient. Defaults to ``GradientStyle.LINEAR``.
            step_count (int, optional): Specifies the number of steps of change color. Defaults to ``0``.
            offset (Offset, int, optional): Specifies the X and Y coordinate, where the gradient begins.
                X is effectively the center of the ``RADIAL``, ``ELLIPTICAL``, ``SQUARE`` and ``RECT``
                style gradients. Defaults to ``Offset(50, 50)``.
            angle (Angle, int, optional): Specifies angle of the gradient. Defaults to 0.
            border (int, optional): Specifies percent of the total width where just the start color is used.
                Defaults to ``0``.
            grad_color (ColorRange, optional): Specifies the color at the start point and stop point of the gradient.
                Defaults to ``ColorRange(Color(0), Color(16777215))``.
            grad_intensity (IntensityRange, optional): Specifies the intensity at the start point and stop point of the
                gradient. Defaults to ``IntensityRange(100, 100)``.
            name (str, optional): Specifies the Fill Gradient Name.
            style_name (StyleParaKind, str, optional): Specifies the Page Style that instance applies to.
                Default is Default Page Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_writer_format_modify_page_header_area`
        """
        direct = InnerGradient(
            style=style,
            step_count=step_count,
            offset=offset,
            angle=angle,
            border=border,
            grad_color=grad_color,
            grad_intensity=grad_intensity,
            name=name,
            _cattribs=self._get_inner_cattribs(),
        )
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())

    # region internal methods
    def _get_inner_props(self) -> AreaGradientProps:
        return AreaGradientProps(
            style="HeaderFillStyle",
            step_count="HeaderFillGradientStepCount",
            name="HeaderFillGradientName",
            grad_prop_name="HeaderFillGradient",
        )

    def _get_inner_cattribs(self) -> dict:
        return {
            "_supported_services_values": self._supported_services(),
            "_format_kind_prop": self.prop_format_kind,
            "_props_internal_attributes": self._get_inner_props(),
        }

    # endregion internal methods

    # region Static Methods
    @classmethod
    def from_style(
        cls: Type[_TGradient],
        doc: Any,
        style_name: WriterStylePageKind | str = WriterStylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> _TGradient:
        """
        Gets instance from Document.

        Args:
            doc (Any): UNO Document Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            Gradient: ``Gradient`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerGradient.from_obj(inst.get_style_props(doc), _cattribs=inst._get_inner_cattribs())
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    @classmethod
    def from_preset(
        cls: Type[_TGradient],
        preset: PresetGradientKind,
        style_name: WriterStylePageKind | str = WriterStylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> _TGradient:
        """
        Gets instance from preset.

        Args:
            preset (PresetGradientKind): Preset.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            Gradient: ``Gradient`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerGradient.from_preset(preset=preset, _cattribs=inst._get_inner_cattribs())
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    # endregion Static Methods

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.DOC | FormatKind.STYLE | FormatKind.HEADER
        return self._format_kind_prop

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | WriterStylePageKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> InnerGradient:
        """Gets/Sets Inner Gradient instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerGradient, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerGradient) -> None:
        if not isinstance(value, InnerGradient):
            raise TypeError(f'Expected type of InnerGradient, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())

    # endregion Properties
