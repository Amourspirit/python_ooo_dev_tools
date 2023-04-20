from __future__ import annotations
from typing import Any, Tuple, cast, overload
import uno
from com.sun.star.lang import XMultiServiceFactory
from com.sun.star.chart2 import XChartDocument

from ooo.dyn.drawing.fill_style import FillStyle
from ooo.dyn.drawing.hatch_style import HatchStyle

from ooodev.exceptions import ex as mEx
from ooodev.format.inner.direct.structs.hatch_struct import HatchStruct
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.preset import preset_hatch as mPreset
from ooodev.format.inner.preset.preset_hatch import PresetHatchKind
from ooodev.format.inner.style_base import StyleMulti
from ooodev.units import UnitMM
from ooodev.units import UnitObj
from ooodev.units.unit_convert import UnitConvert
from ooodev.utils import lo as mLo
from ooodev.utils.color import Color
from ooodev.utils.data_type.angle import Angle as Angle
from ooodev.utils.data_type.intensity import Intensity as Intensity


class Hatch(StyleMulti):
    """
    Class for Chart Area Fill Hatch.

    .. versionadded:: 0.9.4
    """

    def __init__(
        self,
        chart_doc: XChartDocument,
        *,
        style: HatchStyle = HatchStyle.SINGLE,
        color: Color = Color(0),
        space: float | UnitObj = 0,
        angle: Angle | int = 0,
        bg_color: Color = Color(-1),
    ) -> None:
        """
        Constructor.

        Args:
            chart_doc (XChartDocument): Chart document.
            style (HatchStyle, optional): Specifies the kind of lines used to draw this hatch. Default ``HatchStyle.SINGLE``.
            color (:py:data:`~.utils.color.Color`, optional): Specifies the color of the hatch lines. Default ``0``.
            space (float, UnitObj, optional): Specifies the space between the lines in the hatch (in ``mm`` units) or :ref:`proto_unit_obj`. Default ``0.0``
            angle (Angle, int, optional): Specifies angle of the hatch in degrees. Default to ``0``.
            bg_color(Color, optionl): Specifies the background Color. Set this ``-1`` (default) for no background color.

        Returns:
            None:
        """
        super().__init__()

        self._chart_doc = chart_doc
        self._style = style
        self._color = color

        if self._color < 0:
            self._color = 0

        try:
            self._space = space.get_value_mm100()
        except AttributeError:
            self._space = UnitConvert.convert_mm_mm100(space)

        try:
            self._angle = angle.value
        except AttributeError:
            self._angle = int(angle)

        hatch_str_name = getattr(self, "_hatch_str_name", "")
        _ = self._get_hatch(name=hatch_str_name, auto_name=False, allow_update=False)

        self._set("FillStyle", FillStyle.HATCH)
        if bg_color > -1:
            self._set("FillBackground", True)
            self._set("FillColor", bg_color)
        else:
            self._set("FillBackground", False)

    # region Internal Methods
    def _get_hatch(self, name: str, auto_name: bool, allow_update: bool) -> Any:
        # if the name passed in already exist in the hatch Table then it is returned.
        # Otherwise the hatch is added to the hatch Table and then returned.
        # after hatch is added to table all other subsequent call of this name will return
        # that hatch from the Table. With the exception of auto_name which will force a new entry
        # into the Table each time.
        if not name:
            auto_name = True
            name = "ChartHatch "

        if PresetHatchKind.is_preset(name):
            # preset hatch does not need to be added to the hatch table
            self._set("FillHatchName", name)
            return None

        nc = self._container_get_inst()
        if auto_name:
            hatch_name = self._container_get_unique_el_name(name, nc)
        else:
            hatch_name = name
        c_value = self._container_get_value(hatch_name, nc)  # raises value error if name is empty
        if not self.prop_hatch_name:
            self._set("FillHatchName", hatch_name)
        if not c_value is None:
            return c_value
        hatch_obj = self._get_inner_class()
        self._container_add_value(name=hatch_name, obj=hatch_obj.get_uno_struct(), allow_update=allow_update, nc=nc)
        return self._container_get_value(hatch_name, nc)

    def _get_inner_class(self) -> HatchStruct:
        hatch = HatchStruct(
            style=self._style,
            color=self._color,
            distance=self._space,
            angle=self._angle,
        )
        return hatch

    def _update_container_hatch(self) -> None:
        _ = self._get_hatch(name=self._name, auto_name=False, allow_update=True)

    # endregion Internal Methods

    # region Overridden Methods
    # region copy()
    @overload
    def copy(self) -> Hatch:
        ...

    @overload
    def copy(self, **kwargs) -> Hatch:
        ...

    def copy(self, **kwargs) -> Hatch:
        """
        Copy the current instance.

        Returns:
            Hatch: The copied instance.
        """
        cp = super().copy(**kwargs)
        cp._angle = self._angle
        cp._chart_doc = self._chart_doc
        cp._color = self._color
        cp._space = self._space
        cp._style = self._style
        return cp

    # endregion copy()

    def _container_get_service_name(self) -> str:
        # https://github.com/LibreOffice/core/blob/d9e044f04ac11b76b9a3dac575f4e9155b67490e/chart2/source/tools/PropertyHelper.cxx#L229
        return "com.sun.star.drawing.HatchTable"

    def _container_get_msf(self) -> XMultiServiceFactory | None:
        if self._chart_doc is not None:
            chart_doc_ms_factory = mLo.Lo.qi(XMultiServiceFactory, self._chart_doc)
            return chart_doc_ms_factory
        return None

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

    def _props_set(self, obj: object, **kwargs: Any) -> None:
        try:
            super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"Hatch.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion Overridden Methods

    # region Static Methods
    @classmethod
    def from_preset(cls, chart_doc: XChartDocument, preset: PresetHatchKind, **kwargs) -> Hatch:
        """
        Gets an instance from a preset.

        Args:
            chart_doc (XChartDocument): Chart document.
            preset (~.format.inner.preset.preset_hatch.PresetHatchKind): Preset.

        Returns:
            Hatch: Instance from preset.
        """
        cattribs = kwargs.pop("_cattribs", {})
        if not "_hatch_str_name" in cattribs:
            cattribs["_hatch_str_name"] = str(preset)
        kargs = mPreset.get_preset(preset)
        kargs.update(kwargs)
        return cls(chart_doc=chart_doc, _cattribs=cattribs, **kargs)

    # endregion Static Methods

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.UNKNOWN
        return self._format_kind_prop

    @property
    def prop_angle(self) -> Angle:
        """Gets/Sets the angle."""
        return Angle(self._angle)

    @prop_angle.setter
    def prop_angle(self, value: Angle | int) -> None:
        """Sets the angle."""
        try:
            self._angle = value.value
        except AttributeError:
            self._angle = int(value)
        self._update_container_hatch()

    @property
    def prop_bg_color(self) -> Color:
        """Gets/Sets the background color."""
        pv = cast(int, self._get("FillColor"))
        if pv is None:
            return -1
        return pv

    @prop_bg_color.setter
    def prop_bg_color(self, value: Color) -> None:
        """Sets the background color."""
        if value < 0:
            self._set("FillBackground", False)
            self._remove("FillColor")
        else:
            self._set("FillBackground", True)
            self.set("FillColor", value)

    @property
    def prop_color(self) -> Color:
        """Gets/Sets the color."""
        return self._color

    @prop_color.setter
    def prop_color(self, value: Color) -> None:
        """Sets the color."""
        if value < 0:
            value = 0
        self._color = value
        self._update_container_hatch()

    @property
    def prop_style(self) -> HatchStyle:
        """Gets/Sets the style."""
        return self._style

    @prop_style.setter
    def prop_style(self, value: HatchStyle) -> None:
        """Sets the style."""
        self._style = value
        self._update_container_hatch()

    @property
    def prop_space(self) -> UnitMM:
        """Gets/Sets the space."""
        return UnitMM.from_mm100(self._space)

    @property
    def prop_hatch_name(self) -> str:
        """Gets Hatch Name."""
        pv = cast(str, self._get("FillHatchName"))
        if pv is None:
            return ""
        return pv

    @prop_space.setter
    def prop_space(self, value: float | UnitObj) -> None:
        """Sets the space."""
        try:
            self._space = value.get_value_mm100()
        except AttributeError:
            self._space = UnitConvert.convert_mm_mm100(value)
        self._update_container_hatch()

    # endregion Properties