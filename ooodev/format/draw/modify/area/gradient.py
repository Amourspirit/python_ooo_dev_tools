"""Module for Draw Style Fill Coloring."""

# region Import
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from ooo.dyn.awt.gradient_style import GradientStyle

from ooodev.format.draw.style.kind import DrawStyleFamilyKind
from ooodev.format.draw.style.lookup import FamilyGraphics
from ooodev.format.inner.direct.write.fill.area.gradient import Gradient as InnerGradient
from ooodev.format.inner.modify.write.fill.fill_style_base_multi import FillStyleBaseMulti
from ooodev.format.inner.preset.preset_gradient import PresetGradientKind
from ooodev.format.writer.style.para.kind.style_para_kind import StyleParaKind as StyleParaKind
from ooodev.utils.color import Color
from ooodev.utils.data_type.color_range import ColorRange
from ooodev.utils.data_type.intensity_range import IntensityRange
from ooodev.utils.data_type.offset import Offset

if TYPE_CHECKING:
    from ooodev.units.angle import Angle
    from ooodev.utils.data_type.intensity import Intensity

# endregion Import


class Gradient(FillStyleBaseMulti):
    """
    Draw Style Fill Gradient

    .. seealso::

        - :ref:`help_draw_format_modify_area_gradient`

    .. versionadded:: 0.17.9
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
        style_name: str = FamilyGraphics.DEFAULT_DRAWING_STYLE,
        style_family: str | DrawStyleFamilyKind = DrawStyleFamilyKind.GRAPHICS,
    ) -> None:
        """
        Constructor

        Args:
            style (GradientStyle, optional): Specifies the style of the gradient. Defaults to ``GradientStyle.LINEAR``.
            step_count (int, optional): Specifies the number of steps of change color. Defaults to ``0``.
            offset (Offset, int, optional): Specifies the X and Y coordinate, where the gradient begins.
                X is effectively the center of the ``RADIAL``, ``ELLIPTICAL``, ``SQUARE`` and ``RECT``
                style gradients. Defaults to ``Offset(50, 50)``.
            angle (Angle, int, optional): Specifies angle of the gradient. Defaults to ``0``.
            border (int, optional): Specifies percent of the total width where just the start color is used.
                Defaults to ``0``.
            grad_color (ColorRange, optional): Specifies the color at the start point and stop point of the gradient.
                Defaults to ``ColorRange(Color(0), Color(16777215))``.
            grad_intensity (IntensityRange, optional): Specifies the intensity at the start point and stop point of the
                gradient. Defaults to ``IntensityRange(100, 100)``.
            name (str, optional): Specifies the Fill Gradient Name.
            style_name (FamilyGraphics, str, optional): Specifies the Style that instance applies to.
                Default is Default ``standard`` Style.
            style_family (str, DrawStyleFamilyKind, optional): Family Style. Defaults to ``graphics``.

        Returns:
            None:

        See Also:
            - :ref:`help_draw_format_modify_area_gradient`
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
        )
        super().__init__()
        self._style_family_name = str(style_family)
        self._style_name = str(style_name)
        self._set_style("direct", direct)

    @classmethod
    def from_preset(
        cls,
        preset: PresetGradientKind,
        style_name: str = FamilyGraphics.DEFAULT_DRAWING_STYLE,
        style_family: str | DrawStyleFamilyKind = DrawStyleFamilyKind.GRAPHICS,
    ) -> Gradient:
        """
        Gets an instance from a preset

        Args:
            preset (PresetGradientKind): Preset

        Returns:
            Gradient: Instance from preset.
        """
        direct = InnerGradient.from_preset(preset)
        inst = cls(style_name=style_name, style_family=style_family)
        inst._set_style("direct", direct)
        return inst

    @classmethod
    def from_style(
        cls,
        doc: Any,
        style_name: FamilyGraphics | str = FamilyGraphics.DEFAULT_DRAWING_STYLE,
        style_family: str | DrawStyleFamilyKind = DrawStyleFamilyKind.GRAPHICS,
    ) -> Gradient:
        """
        Gets instance from Document.

        Args:
            doc (Any): UNO Document Object.
            style_name (FamilyGraphics, str, optional): Specifies the Style that instance applies to.
                Default is ``FamilyGraphics.DEFAULT_DRAWING_STYLE``.
            style_family (DrawStyleFamilyKind, str, optional): Style family. Default ``DrawStyleFamilyKind.GRAPHICS``.

        Returns:
            Gradient: ``Gradient`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerGradient.from_obj(inst.get_style_props(doc))
        inst._set_style("direct", direct)
        return inst

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | FamilyGraphics):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> InnerGradient:
        """Gets Inner instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerGradient, self._get_style_inst("direct"))
        return self._direct_inner
