from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING
import uno

from ooodev.mock.mock_g import DOCS_BUILDING
from ooodev.format.proto.style_t import StyleT

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Protocol
    from ooo.dyn.awt.gradient import Gradient
    from ooo.dyn.awt.gradient_style import GradientStyle
    from ooodev.format.inner.preset.preset_gradient import PresetGradientKind
    from ooodev.utils.color import Color
    from ooodev.units.angle import Angle
    from ooodev.utils.data_type.intensity import Intensity
else:
    Protocol = object
    Gradient = Any
    GradientStyle = Any
    PresetGradientKind = Any
    Color = Any
    Angle = Any
    Intensity = Any


class GradientStructT(StyleT, Protocol):
    """Gradient Struct Protocol"""

    def __init__(
        self,
        *,
        style: GradientStyle = ...,
        step_count: int = ...,
        x_offset: Intensity | int = ...,
        y_offset: Intensity | int = ...,
        angle: Angle | int = ...,
        border: Intensity | int = ...,
        start_color: Color = ...,
        start_intensity: Intensity | int = ...,
        end_color: Color = ...,
        end_intensity: Intensity | int = ...,
    ) -> None: ...

    def get_uno_struct(self) -> Gradient:
        """
        Gets UNO ``Gradient`` from instance.

        Returns:
            Gradient: ``Gradient`` instance
        """
        ...

    def get_json(self) -> str:
        """
        Get Gradient represented as a json string for use with dispatch commands.

        Returns:
            str: Json string.
        """
        ...

    @overload
    @classmethod
    def from_uno_struct(cls, value: Gradient) -> GradientStructT: ...

    @overload
    @classmethod
    def from_uno_struct(cls, value: Gradient, **kwargs) -> GradientStructT: ...

    @classmethod
    def from_uno_struct(cls, value: Gradient, **kwargs) -> GradientStructT: ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> GradientStructT: ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> GradientStructT: ...

    @overload
    @classmethod
    def from_preset(cls, preset: PresetGradientKind) -> GradientStructT: ...

    @overload
    @classmethod
    def from_preset(cls, preset: PresetGradientKind, **kwargs) -> GradientStructT: ...

    @classmethod
    def from_preset(cls, preset: PresetGradientKind, **kwargs) -> GradientStructT: ...

    # region Style Properties

    # endregion Style Properties

    # region Prop Properties
    @property
    def prop_style(self) -> GradientStyle:
        """Gets/Sets the style of the gradient."""
        ...

    @prop_style.setter
    def prop_style(self, value: GradientStyle): ...

    @property
    def prop_step_count(self) -> int:
        """Gets/Sets the number of steps of change color."""
        ...

    @prop_step_count.setter
    def prop_step_count(self, value: int): ...

    @property
    def prop_x_offset(self) -> Intensity:
        """Gets/Sets the X-coordinate, where the gradient begins."""
        ...

    @prop_x_offset.setter
    def prop_x_offset(self, value: Intensity | int): ...

    @property
    def prop_y_offset(self) -> Intensity:
        """Gets/Sets the Y-coordinate, where the gradient begins."""
        ...

    @prop_y_offset.setter
    def prop_y_offset(self, value: Intensity | int): ...

    @property
    def prop_angle(self) -> Angle:
        """Gets/Sets angle of the gradient."""
        ...

    @prop_angle.setter
    def prop_angle(self, value: Angle | int): ...

    @property
    def prop_border(self) -> Intensity:
        """Gets/Sets percent of the total width where just the start color is used."""
        ...

    @prop_border.setter
    def prop_border(self, value: Intensity | int): ...

    @property
    def prop_start_color(self) -> Color:
        """Gets/Sets the color at the start point of the gradient."""
        ...

    @prop_start_color.setter
    def prop_start_color(self, value: Color): ...

    @property
    def prop_start_intensity(self) -> Intensity:
        """Gets/Sets the intensity at the start point of the gradient."""
        ...

    @prop_start_intensity.setter
    def prop_start_intensity(self, value: Intensity | int): ...

    @property
    def prop_end_color(self) -> Color:
        """Gets/Sets the color at the end point of the gradient."""
        ...

    @prop_end_color.setter
    def prop_end_color(self, value: Color): ...

    @property
    def prop_end_intensity(self) -> Intensity:
        """Gets/Sets the intensity at the end point of the gradient."""
        ...

    @prop_end_intensity.setter
    def prop_end_intensity(self, value: Intensity | int): ...

    # endregion Prop Properties
