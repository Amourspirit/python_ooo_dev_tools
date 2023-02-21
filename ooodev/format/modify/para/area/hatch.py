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
from ....preset import preset_hatch
from ....preset.preset_hatch import PresetHatchKind as PresetHatchKind
from ..para_style_base_multi import ParaStyleBaseMulti

from ....direct.para.area.hatch import Hatch as DirectHatch

# from ....direct.fill.area.hatch import Hatch as DirectHatch

from ooo.dyn.drawing.hatch_style import HatchStyle as HatchStyle


class Hatch(ParaStyleBaseMulti):
    """
    Paragraph Style Gradient Color

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        style: HatchStyle = HatchStyle.SINGLE,
        color: Color = Color(0),
        space: float = 0.0,
        angle: Angle | int = 0,
        bg_color: Color = Color(-1),
        name: str = "",
        auto_name: bool = False,
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

        direct = DirectHatch(
            style=style, color=color, space=space, angle=angle, bg_color=bg_color, name=name, auto_name=auto_name
        )
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())

    @staticmethod
    def from_preset(
        preset: PresetHatchKind,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> Hatch:
        """
        Gets instance from preset

        Args:
            preset (PresetHatchKind): Preset

        Returns:
            Hatch: Hatch from a preset.
        """
        args = preset_hatch.get_preset(preset)
        args["name"] = str(preset)
        args["style_name"] = style_name
        args["style_family"] = style_family
        return Hatch(**args)

    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> Hatch:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Documnet Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``ParagraphStyles``.

        Returns:
            Alignment: ``Alignment`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = DirectHatch.from_obj(inst.get_style_props(doc))
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
    def prop_inner(self) -> DirectHatch:
        """Gets Inner Hatch instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(DirectHatch, self._get_style_inst("direct"))
        return self._direct_inner
