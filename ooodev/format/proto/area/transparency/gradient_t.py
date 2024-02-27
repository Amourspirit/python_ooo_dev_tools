from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING
import uno
from ooodev.mock.mock_g import DOCS_BUILDING

from ooodev.format.proto.style_multi_t import StyleMultiT

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Protocol
    from ooo.dyn.awt.gradient_style import GradientStyle

    from ooodev.format.inner.direct.structs.gradient_struct import GradientStruct
    from ooodev.units.angle import Angle
    from ooodev.utils.data_type.intensity import Intensity
    from ooodev.utils.data_type.intensity_range import IntensityRange
    from ooodev.utils.data_type.offset import Offset
else:
    Protocol = object
    GradientStyle = Any
    GradientStruct = Any
    Angle = Any
    Intensity = None
    IntensityRange = Any
    Offset = Any


class GradientT(StyleMultiT, Protocol):
    """Fill Area Transparency Gradient Protocol"""

    def __init__(
        self,
        *,
        style: GradientStyle = ...,
        offset: Offset = ...,
        angle: Angle | int = ...,
        border: Intensity | int = ...,
        grad_intensity: IntensityRange = ...,
        **kwargs: Any,
    ) -> None: ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> GradientT: ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> GradientT: ...

    # region Properties
    @property
    def prop_inner(self) -> GradientStruct:
        """Gets Fill Transparent Gradient instance"""
        ...

    @property
    def default(self) -> GradientT:
        """Gets Gradient empty. Static Property."""
        ...

    # endregion Properties
