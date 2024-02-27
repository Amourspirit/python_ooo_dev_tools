from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING
import uno

from ooodev.mock.mock_g import DOCS_BUILDING
from ooodev.format.proto.style_multi_t import StyleMultiT

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Self
    from typing_extensions import Protocol
    from ooo.dyn.awt.gradient_style import GradientStyle
    from ooodev.events.args.key_val_cancel_args import KeyValCancelArgs
    from ooodev.units.angle import Angle
    from ooodev.utils.data_type.intensity import Intensity
    from ooodev.utils.data_type.intensity_range import IntensityRange
    from ooodev.utils.data_type.offset import Offset as Offset
    from ooodev.format.inner.direct.structs.gradient_struct import GradientStruct
else:
    Protocol = object
    Self = Any
    GradientStyle = Any
    KeyValCancelArgs = Any
    Angle = Any
    Intensity = Any
    IntensityRange = Any
    Offset = Any
    GradientStruct = Any


class FillGradientT(StyleMultiT, Protocol):
    """Fill Gradient Protocol"""

    def __init__(
        self,
        *,
        style: GradientStyle = ...,
        offset: Offset = ...,
        angle: Angle | int = ...,
        border: Intensity | int = ...,
        grad_intensity: IntensityRange = ...,
        **kwargs: Any,
    ) -> None:
        """
        Constructor

        Args:
            style (GradientStyle, optional): Specifies the style of the gradient. Defaults to ``GradientStyle.LINEAR``.
            step_count (int, optional): Specifies the number of steps of change color. Defaults to ``0``.
            offset (offset, optional): Specifies the X-coordinate (start) and Y-coordinate (end),
                where the gradient begins. X is effectively the center of the ``RADIAL``, ``ELLIPTICAL``, ``SQUARE`` and
                ``RECT`` style gradients. Defaults to ``Offset(50, 50)``.
            angle (Angle, int, optional): Specifies angle of the gradient. Defaults to ``0``.
            border (int, optional): Specifies percent of the total width where just the start color is used.
                Defaults to ``0``.
            grad_intensity (IntensityRange, optional): Specifies the intensity at the start point and stop point of
                the gradient. Defaults to ``IntensityRange(0, 0)``.

        Returns:
            None:
        """
        ...

    def on_property_setting(self, source: Any, event_args: KeyValCancelArgs) -> None: ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> Self: ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> Self: ...

    @property
    def prop_inner(self) -> GradientStruct:
        """Gets Fill Transparent Gradient instance"""
        ...

    @property
    def default(self) -> Self:
        """Gets Gradient empty. Static Property."""
        ...
