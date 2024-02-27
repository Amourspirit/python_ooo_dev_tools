"""
Module for Fill Gradient Color.

.. versionadded:: 0.9.0
"""

# region Import
from __future__ import annotations
from typing import Any, Tuple, overload, TYPE_CHECKING
import uno
from com.sun.star.lang import XMultiServiceFactory
from com.sun.star.chart2 import XChartDocument

from ooo.dyn.awt.gradient_style import GradientStyle

from ooodev.loader import lo as mLo
from ooodev.format.inner.direct.structs.gradient_struct import GradientStruct
from ooodev.format.inner.direct.write.fill.transparent.gradient import Gradient as WriteGradient
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.units.angle import Angle
from ooodev.utils.data_type.intensity import Intensity
from ooodev.utils.data_type.intensity_range import IntensityRange
from ooodev.utils.data_type.offset import Offset
from ooodev.meta.disabled_method import DisabledMethod
from ooodev.meta.deleted_attrib import DeletedAttrib

if TYPE_CHECKING:
    from ooo.dyn.awt.gradient import Gradient as UNOGradient

# endregion Import


class _GradientStruct(GradientStruct):
    def _get_property_name(self) -> str:
        return ""


class Gradient(WriteGradient):
    """
    Chart Fill Gradient Color

    .. seealso::

        - :ref:`help_chart2_format_direct_general_transparency`

    .. versionadded:: 0.9.4
    """

    from_obj = DisabledMethod()  # type: ignore
    default = DeletedAttrib()  # type: ignore

    def __init__(
        self,
        chart_doc: XChartDocument,
        *,
        style: GradientStyle = GradientStyle.LINEAR,
        offset: Offset = Offset(50, 50),
        angle: Angle | int = 0,
        border: Intensity | int = 0,
        grad_intensity: IntensityRange = IntensityRange(0, 0),
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

        See Also:
            - :ref:`help_chart2_format_direct_general_transparency`
        """
        self._chart_doc = chart_doc
        super().__init__(
            style=style, offset=offset, angle=angle, border=border, grad_intensity=grad_intensity, **kwargs
        )

    # region overrides
    def _container_get_msf(self) -> XMultiServiceFactory | None:
        if self._chart_doc is not None:
            return mLo.Lo.qi(XMultiServiceFactory, self._chart_doc)
        return None

    def _container_get_default_name(self) -> str:
        return "ChartTransparencyGradient"

    def _get_gradient_from_uno_struct(self, value: UNOGradient, **kwargs) -> GradientStruct:
        return _GradientStruct.from_uno_struct(value, **kwargs)

    def _get_inner_class(
        self,
        style: GradientStyle,
        step_count: int,
        x_offset: Intensity | int,
        y_offset: Intensity | int,
        angle: Angle | int,
        border: Intensity | int,
        start_color: int,
        start_intensity: Intensity | int,
        end_color: int,
        end_intensity: Intensity | int,
    ) -> _GradientStruct:
        # pylint: disable=unexpected-keyword-arg
        return _GradientStruct(
            style=style,
            step_count=step_count,
            x_offset=x_offset,
            y_offset=y_offset,
            angle=angle,
            border=border,
            start_color=start_color,
            start_intensity=start_intensity,
            end_color=end_color,
            end_intensity=end_intensity,
            _cattribs=self._get_inner_cattribs(),
        )

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.chart2.PageBackground",
                "com.sun.star.chart2.DataSeries",
                "com.sun.star.chart2.DataPoint",
                "com.sun.star.chart2.Legend",
            )
        return self._supported_services_values

    def _is_valid_obj(self, obj: Any) -> bool:
        return self._is_obj_service(obj)

    # region copy()
    @overload
    def copy(self) -> Gradient: ...

    @overload
    def copy(self, **kwargs) -> Gradient: ...

    def copy(self, **kwargs) -> Gradient:
        """
        Copy the current instance.

        Returns:
            Hatch: The copied instance.
        """
        # pylint: disable=protected-access
        cp = super().copy(**kwargs)
        cp._chart_doc = self._chart_doc
        return cp

    # endregion copy()

    # endregion overrides

    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.UNKNOWN
        return self._format_kind_prop

    # endregion properties
