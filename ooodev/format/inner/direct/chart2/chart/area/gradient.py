# region Import
from __future__ import annotations
from typing import Tuple, overload, TYPE_CHECKING
import uno
from com.sun.star.lang import XMultiServiceFactory
from com.sun.star.chart2 import XChartDocument


from ooo.dyn.awt.gradient_style import GradientStyle as GradientStyle

from ooodev.utils import lo as mLo
from ooodev.utils.color import Color
from ooodev.utils.data_type.angle import Angle as Angle
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
from ooodev.format.inner.direct.structs.gradient_struct import GradientStruct

if TYPE_CHECKING:
    from ooo.dyn.awt.gradient import Gradient as UNOGradient

# endregion Import


class _TitleGradidentStruct(GradientStruct):
    def _get_property_name(self) -> str:
        return ""


class Gradient(FillGradient):
    """
    Style for Chart Area Fill Gradient.

    .. versionadded:: 0.9.4
    """

    default = DeletedAttrib()

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
                "com.sun.star.chart2.PageBackground",
                "com.sun.star.chart2.Title",
                "com.sun.star.chart2.DataSeries",
                "com.sun.star.chart2.DataPoint",
                "com.sun.star.chart2.Legend",
            )
        return self._supported_services_values

    def _container_get_msf(self) -> XMultiServiceFactory | None:
        if self._chart_doc is not None:
            chart_doc_ms_factory = mLo.Lo.qi(XMultiServiceFactory, self._chart_doc)
            return chart_doc_ms_factory
        return None

    # region copy()
    @overload
    def copy(self) -> Gradient:
        ...

    @overload
    def copy(self, **kwargs) -> Gradient:
        ...

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

    def _get_gradient_from_uno_struct(self, uno_struct: UNOGradient, **kwargs) -> _TitleGradidentStruct:
        _TitleGradidentStruct.from_uno_struct(uno_struct, **kwargs)

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
    ) -> _TitleGradidentStruct:
        fs = _TitleGradidentStruct(
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
        return fs

    # endregion overrides

    # region Static Methods
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, chart_doc: XChartDocument, obj: object) -> Gradient:
        ...

    @overload
    @classmethod
    def from_obj(cls, chart_doc: XChartDocument, obj: object, **kwargs) -> Gradient:
        ...

    @classmethod
    def from_obj(cls, chart_doc: XChartDocument, obj: object, **kwargs) -> Gradient:
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
        return super().from_obj(obj=obj, chart_doc=chart_doc, **kwargs)

    # endregion from_obj()

    # region from_preset()
    @overload
    @classmethod
    def from_preset(cls, chart_doc: XChartDocument, preset: PresetGradientKind) -> Gradient:
        ...

    @overload
    @classmethod
    def from_preset(cls, chart_doc: XChartDocument, preset: PresetGradientKind, **kwargs) -> Gradient:
        ...

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
