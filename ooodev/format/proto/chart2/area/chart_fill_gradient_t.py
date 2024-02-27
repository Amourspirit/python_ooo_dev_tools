from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING
import uno


from ooodev.mock.mock_g import DOCS_BUILDING
from ooodev.format.proto.area.fill_gradient_t import FillGradientT

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Protocol
    from com.sun.star.chart2 import XChartDocument
    from ooo.dyn.awt.gradient_style import GradientStyle
    from ooodev.units.angle import Angle
    from ooodev.utils.data_type.color_range import ColorRange
    from ooodev.utils.data_type.intensity import Intensity
    from ooodev.utils.data_type.intensity_range import IntensityRange
    from ooodev.utils.data_type.offset import Offset
    from ooodev.format.inner.preset.preset_gradient import PresetGradientKind
    from ooodev.format.inner.direct.structs.gradient_struct import GradientStruct
else:
    Protocol = object
    XChartDocument = Any
    GradientStyle = Any
    Angle = Any
    ColorRange = Any
    Intensity = Any
    IntensityRange = Any
    Offset = Any
    PresetGradientKind = Any
    GradientStruct = Any


class ChartFillGradientT(FillGradientT, Protocol):
    """Chart Fill Gradient Protocol"""

    def __init__(
        self,
        chart_doc: XChartDocument,
        *,
        style: GradientStyle = ...,
        step_count: int = ...,
        offset: Offset = ...,
        angle: Angle | int = ...,
        border: Intensity | int = ...,
        grad_color: ColorRange = ...,
        grad_intensity: IntensityRange = ...,
        name: str = ...,
    ) -> None:
        """
        Constructor

        Args:
            chart_doc (XChartDocument): Chart document.
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

        Returns:
            None:

        See Also:
            - :ref:`help_chart2_format_direct_general_area`
        """
        ...

    @overload
    @classmethod
    def from_obj(cls, chart_doc: XChartDocument, obj: Any) -> ChartFillGradientT: ...

    @overload
    @classmethod
    def from_obj(cls, chart_doc: XChartDocument, obj: Any, **kwargs) -> ChartFillGradientT: ...

    @overload
    @classmethod
    def from_preset(cls, chart_doc: XChartDocument, preset: PresetGradientKind) -> ChartFillGradientT: ...

    @overload
    @classmethod
    def from_preset(cls, chart_doc: XChartDocument, preset: PresetGradientKind, **kwargs) -> ChartFillGradientT: ...

    @classmethod
    def from_struct(
        cls, chart_doc: XChartDocument, struct: GradientStruct, name: str = "", **kwargs
    ) -> ChartFillGradientT:
        """
        Gets instance from ``GradientStruct`` instance

        Args:
            chart_doc (XChartDocument): Chart document.
            struct (GradientStruct): Gradient Struct instance.
            name (str, optional): Name of Gradient.

        Returns:
            Gradient:
        """
        ...

    @property
    def default(self) -> ChartFillGradientT:
        """Gets Gradient empty."""
        ...
