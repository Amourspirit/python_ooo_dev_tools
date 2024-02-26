# region Import
from __future__ import annotations
from typing import Any, cast, Tuple, overload, TYPE_CHECKING
import uno
from com.sun.star.lang import XMultiServiceFactory
from com.sun.star.chart2 import XChartDocument


from ooo.dyn.awt.gradient_style import GradientStyle as GradientStyle

from ooodev.loader import lo as mLo
from ooodev.utils.color import Color
from ooodev.units.angle import Angle as Angle
from ooodev.utils.data_type.color_range import ColorRange as ColorRange
from ooodev.utils.data_type.intensity import Intensity as Intensity
from ooodev.utils.data_type.intensity_range import IntensityRange as IntensityRange
from ooodev.utils.data_type.offset import Offset as Offset
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.preset.preset_gradient import (
    PresetGradientKind as PresetGradientKind,
)
from ooodev.format.inner.direct.structs.gradient_struct import GradientStruct
from ooodev.format.inner.direct.write.fill.area.gradient import Gradient as FillGradient
from ooodev.meta.deleted_attrib import DeletedAttrib
from ooodev.exceptions import ex as mEx
from ooodev.utils import props as mProps

if TYPE_CHECKING:
    from ooo.dyn.awt.gradient import Gradient as UNOGradient

# endregion Import


class _TitleGradientStruct(GradientStruct):
    def _get_property_name(self) -> str:
        return ""


class Gradient(FillGradient):
    """
    Style for Chart Area Fill Gradient.

    .. seealso::

        - :ref:`help_chart2_format_direct_general_area`

    .. versionadded:: 0.9.4
    """

    default = DeletedAttrib()  # type: ignore

    def __init__(
        self,
        chart_doc: XChartDocument,
        *,
        style: GradientStyle = GradientStyle.LINEAR,
        step_count: int = 0,
        offset: Offset = Offset(50, 50),
        angle: Angle | int = 0,
        border: Intensity | int = 0,
        grad_color: ColorRange = ColorRange(Color(0), Color(16777215)),
        grad_intensity: IntensityRange = IntensityRange(100, 100),
        name: str = "",
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
        self._chart_doc = chart_doc
        super().__init__(
            style=style,
            step_count=step_count,
            offset=offset,
            angle=angle,
            border=border,
            grad_color=grad_color,
            grad_intensity=grad_intensity,
            name=name,
        )

    # region overrides

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.chart2.DataPoint",
                "com.sun.star.chart2.DataSeries",
                "com.sun.star.chart2.Legend",
                "com.sun.star.chart2.PageBackground",
                "com.sun.star.chart2.Title",
                "com.sun.star.drawing.FillProperties",
            )
        return self._supported_services_values

    def _container_get_msf(self) -> XMultiServiceFactory | None:
        if self._chart_doc is not None:
            return mLo.Lo.qi(XMultiServiceFactory, self._chart_doc)
        return None

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
        cp = super().copy(**kwargs)
        cp._chart_doc = self._chart_doc
        return cp

    # endregion copy()

    def _get_gradient_from_uno_struct(self, uno_struct: UNOGradient, **kwargs) -> _TitleGradientStruct:
        return _TitleGradientStruct.from_uno_struct(uno_struct, **kwargs)

    def _get_inner_class(
        self,
        style: GradientStyle,
        step_count: int,
        x_offset: Intensity | int,
        y_offset: Intensity | int,
        angle: Angle | int,
        border: Intensity | int,
        start_color: Color,
        start_intensity: Intensity | int,
        end_color: Color,
        end_intensity: Intensity | int,
    ) -> _TitleGradientStruct:
        # pylint: disable=unexpected-keyword-arg
        return _TitleGradientStruct(
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
            _cattribs=self._get_gradient_struct_cattrib(),
        )

    # endregion overrides

    # region Static Methods
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, chart_doc: XChartDocument, obj: Any) -> Gradient: ...

    @overload
    @classmethod
    def from_obj(cls, chart_doc: XChartDocument, obj: Any, **kwargs) -> Gradient: ...

    @classmethod
    def from_obj(cls, chart_doc: XChartDocument, obj: Any, **kwargs) -> Gradient:
        """
        Gets instance from object.

        Args:
            chart_doc (XChartDocument): Chart document.
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Gradient: Instance that represents Gradient color.
        """
        # pylint: disable=protected-access
        # return super().from_obj(obj=obj, chart_doc=chart_doc, **kwargs)
        inst = cls(name="__constructor_default__", chart_doc=chart_doc, **kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        step_count = cast(int, mProps.Props.get(obj, inst._props.step_count))
        name = cast(str, mProps.Props.get(obj, inst._props.name))
        result = cls(name=name, step_count=step_count, chart_doc=chart_doc, **kwargs)
        result.set_update_obj(obj)
        return result

    # endregion from_obj()

    # region from_preset()
    @overload
    @classmethod
    def from_preset(cls, chart_doc: XChartDocument, preset: PresetGradientKind) -> Gradient: ...

    @overload
    @classmethod
    def from_preset(cls, chart_doc: XChartDocument, preset: PresetGradientKind, **kwargs) -> Gradient: ...

    @classmethod
    def from_preset(cls, chart_doc: XChartDocument, preset: PresetGradientKind, **kwargs) -> Gradient:
        """
        Gets instance from preset.

        Args:
            chart_doc (XChartDocument): Chart document.
            preset (PresetGradientKind): Preset.

        Returns:
            Gradient: Gradient from a preset.
        """
        return super().from_preset(preset=preset, chart_doc=chart_doc, **kwargs)

    # endregion from_preset()

    # region from_gradient()
    @classmethod
    def from_struct(cls, chart_doc: XChartDocument, struct: GradientStruct, name: str = "", **kwargs) -> Gradient:
        """
        Gets instance from ``GradientStruct`` instance

        Args:
            chart_doc (XChartDocument): Chart document.
            struct (GradientStruct): Gradient Struct instance.
            name (str, optional): Name of Gradient.

        Returns:
            Gradient:
        """
        return super().from_struct(struct=struct, name=name, chart_doc=chart_doc, **kwargs)

    # endregion from_gradient()

    # endregion Static Methods

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.FILL
        return self._format_kind_prop

    # endregion Properties
