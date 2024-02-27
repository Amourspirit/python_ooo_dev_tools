# region Imports
from __future__ import annotations
from typing import cast
import uno
from ooo.dyn.awt.gradient_style import GradientStyle

from ooodev.format.inner.direct.write.fill.area.gradient import Gradient as InnerGradient
from ooodev.format.inner.modify.write.frame.frame_style_base_multi import FrameStyleBaseMulti
from ooodev.format.inner.preset.preset_gradient import PresetGradientKind
from ooodev.format.writer.style.frame.style_frame_kind import StyleFrameKind
from ooodev.units.angle import Angle
from ooodev.utils.color import Color
from ooodev.utils.data_type.color_range import ColorRange
from ooodev.utils.data_type.intensity import Intensity
from ooodev.utils.data_type.intensity_range import IntensityRange
from ooodev.utils.data_type.offset import Offset

# endregion Imports


class Gradient(FrameStyleBaseMulti):
    """
    Frame Style Area Gradient.

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
        style_name: StyleFrameKind | str = StyleFrameKind.FRAME,
        style_family: str = "FrameStyles",
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
            style_name (StyleFrameKind, str, optional): Specifies the Frame Style that instance applies to. Default is Default Frame Style.
            style_family (str, optional): Style family. Default ``FrameStyles``.

        Returns:
            None:
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
        direct._prop_parent = self
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())

    def _get_inner_cattribs(self) -> dict:
        return {"_supported_services_values": self._supported_services(), "_format_kind_prop": self.prop_format_kind}

    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: StyleFrameKind | str = StyleFrameKind.FRAME,
        style_family: str = "FrameStyles",
    ) -> Gradient:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Document Object.
            style_name (StyleFrameKind, str, optional): Specifies the Frame Style that instance applies to. Default is Default Frame Style.
            style_family (str, optional): Style family. Default ``FrameStyles``.

        Returns:
            Gradient: ``Gradient`` instance from style properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerGradient.from_obj(obj=inst.get_style_props(doc), _cattribs=inst._get_inner_cattribs())
        direct._prop_parent = inst
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    @classmethod
    def from_preset(
        cls,
        preset: PresetGradientKind,
        style_name: StyleFrameKind | str = StyleFrameKind.FRAME,
        style_family: str = "FrameStyles",
    ) -> Gradient:
        """
        Gets instance from preset.

        Args:
            preset (PresetGradientKind): Preset.
            style_name (StyleFrameKind, str, optional): Specifies the Frame Style that instance applies to. Default is Default Frame Style.
            style_family (str, optional): Style family. Default ``FrameStyles``.

        Returns:
            Gradient: ``Gradient`` instance from a preset.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerGradient.from_preset(preset=preset, _cattribs=inst._get_inner_cattribs())
        direct._prop_parent = inst
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | StyleFrameKind):
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
        struct = value.prop_inner
        inst = InnerGradient.from_struct(struct=struct, name=value.prop_name, _cattribs=self._get_inner_cattribs())

        self._del_attribs("_direct_inner")
        self._set_style("direct", inst, *inst.get_attrs())
