from __future__ import annotations
from typing import cast
import uno
from .....utils.color import Color
from .....utils.data_type.angle import Angle as Angle
from .....utils.data_type.color_range import ColorRange as ColorRange
from .....utils.data_type.intensity import Intensity as Intensity
from .....utils.data_type.intensity_range import IntensityRange as IntensityRange
from .....utils.data_type.offset import Offset as Offset
from ....writer.style.para.kind import StyleParaKind as StyleParaKind
from ....preset import preset_gradient
from ....preset.preset_gradient import PresetGradientKind as PresetGradientKind
from ..para_style_base_multi import ParaStyleBaseMulti
from ....direct.fill.area.gradient import Gradient as DirectGradient

from ooo.dyn.awt.gradient_style import GradientStyle as GradientStyle


class Gradient(ParaStyleBaseMulti):
    """
    Paragraph Style Gradient Color

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
        grad_color: ColorRange = ColorRange(Color(0), Color(16777215)),
        grad_intensity: IntensityRange = IntensityRange(100, 100),
        name: str = "",
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> None:
        """
        Constructor

        Args:
            style (GradientStyle, optional): Specifies the style of the gradient. Defaults to ``GradientStyle.LINEAR``.
            step_count (int, optional): Specifies the number of steps of change color. Defaults to ``0``.
            offset (Offset, int, optional): Specifies the X and Y coordinate, where the gradient begins.
                 X is is effectively the center of the ``RADIAL``, ``ELLIPTICAL``, ``SQUARE`` and ``RECT`` style gradients. Defaults to ``Offset(50, 50)``.
            angle (Angle, int, optional): Specifies angle of the gradient. Defaults to 0.
            border (int, optional): Specifies percent of the total width where just the start color is used. Defaults to 0.
            grad_color (ColorRange, optional): Specifies the color at the start point and stop point of the gradient. Defaults to ``ColorRange(Color(0), Color(16777215))``.
            grad_intensity (IntensityRange, optional): Specifies the intensity at the start point and stop point of the gradient. Defaults to ``IntensityRange(100, 100)``.
            name (str, optional): Specifies the Fill Gradient Name.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``ParagraphStyles``.

        Returns:
            None:
        """

        direct = DirectGradient(
            style=style,
            step_count=step_count,
            offset=offset,
            angle=angle,
            border=border,
            grad_color=grad_color,
            grad_intensity=grad_intensity,
            name=name,
        )
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())

    @staticmethod
    def from_preset(
        preset: PresetGradientKind,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> Gradient:
        """
        Gets instance from preset

        Args:
            preset (PresetKind): Preset

        Returns:
            Gradient: Graident from a preset.
        """
        args = preset_gradient.get_preset(preset)
        args["style_name"] = style_name
        args["style_family"] = style_family
        return Gradient(**args)

    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> Gradient:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Documnet Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``ParagraphStyles``.

        Returns:
            Alignment: ``Alignment`` instance from document properties.
        """
        inst = super(Gradient, cls).__new__(cls)
        inst.__init__(style_name=style_name, style_family=style_family)
        direct = DirectGradient.from_obj(inst.get_style_props(doc))
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
    def prop_inner(self) -> DirectGradient:
        """Gets Inner Gradient instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(DirectGradient, self._get_style_inst("direct"))
        return self._direct_inner
