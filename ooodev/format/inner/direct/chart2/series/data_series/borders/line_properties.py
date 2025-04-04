from __future__ import annotations
from typing import Any, Tuple, cast, overload, NamedTuple, TypeVar, Type
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.preset.preset_border_line import BorderLineKind, get_preset_series_border_line_props
from ooodev.format.inner.style_base import StyleBase
from ooodev.units.unit_convert import UnitConvert
from ooodev.units.unit_mm import UnitMM
from ooodev.units.unit_obj import UnitT
from ooodev.utils import props as mProps
from ooodev.utils.color import Color
from ooodev.utils.data_type.intensity import Intensity

_TLineProperties = TypeVar("_TLineProperties", bound="LineProperties")


class _LinePropertiesProps(NamedTuple):
    color1: str
    color2: str
    width: str
    transparency1: str
    transparency2: str


class LineProperties(StyleBase):
    """
    This class represents the line properties of a chart data series borders line properties.

    .. seealso::

        - :ref:`help_chart2_format_direct_series_series_borders`
    """

    def __init__(
        self,
        style: BorderLineKind = BorderLineKind.CONTINUOUS,
        color: Color = Color(0),
        width: float | UnitT = 0,
        transparency: int | Intensity = 0,
    ) -> None:
        """
        Constructor.

        Args:
            style (BorderLineKind): Line style. Defaults to ``BorderLineKind.CONTINUOUS``.
            color (Color, optional): Line Color. Defaults to ``Color(0)``.
            width (float, UnitT, optional): Line Width (in ``mm`` units) or :ref:`proto_unit_obj`. Defaults to ``0``.
            transparency (int, Intensity, optional): Line transparency from ``0`` to ``100``. Defaults to ``0``.

        Returns:
            None:

        See Also:
            - :ref:`help_chart2_format_direct_series_series_borders`
        """
        super().__init__()
        self._prop_style = style
        self._set_line_style(style)
        self.prop_color = color
        self.prop_width = width
        self.prop_transparency = transparency

    # region Private Methods

    def _set_line_style(self, style: BorderLineKind):
        props = get_preset_series_border_line_props(kind=style)
        self._set("BorderDashName", props.border_name)
        self._set("BorderStyle", props.border_style)
        self._set("LineCap", props.line_cap)
        self._set("LineDash", props.line_dash)
        self._set("LineDashName", props.line_dash_name)
        self._set("LineStyle", props.line_style)

    # endregion Private Methods

    # region Overridden Methods
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.chart2.DataSeries", "com.sun.star.chart2.DataPoint")
        return self._supported_services_values

    # region copy()
    @overload
    def copy(self) -> LineProperties: ...

    @overload
    def copy(self, **kwargs) -> LineProperties: ...

    def copy(self, **kwargs) -> LineProperties:
        """
        Copy the current instance.

        Returns:
            LineProperties: The copied instance.
        """
        cp = super().copy(**kwargs)
        cp._prop_style = self._prop_style
        return cp

    # endregion copy()
    # endregion Overridden Methods

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TLineProperties], obj: object) -> _TLineProperties: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TLineProperties], obj: object, **kwargs) -> _TLineProperties: ...

    @classmethod
    def from_obj(cls: Type[_TLineProperties], obj: Any, **kwargs) -> _TLineProperties:
        """
        Creates a new instance from ``obj``.

        Args:
            obj (Any): UNO Shape object.

        Returns:
            LineProperties: New instance.
        """
        inst = cls(**kwargs)
        # pylint: disable=protected-access
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError("Object is not supported for conversion to Line Properties")

        props = {
            "BorderDashName",
            "BorderStyle",
            "LineCap",
            "LineDash",
            "LineDashName",
            "LineStyle",
        }

        def set_property(prop: str):
            value = mProps.Props.get(obj, prop, None)
            if value is not None:
                inst._set(prop, value)

        for prop in props:
            set_property(prop)
        inst.set_update_obj(obj)
        return inst

    # endregion from_obj()

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
    def prop_color(self) -> Color:
        """Gets/Sets the color."""
        return self._get(self._props.color1)

    @prop_color.setter
    def prop_color(self, value: Color):
        value = max(value, 0)  # type: ignore
        self._set(self._props.color1, value)
        if self._props.color2:
            self._set(self._props.color2, value)

    @property
    def prop_width(self) -> UnitMM:
        pv = cast(int, self._get(self._props.width))
        return UnitMM.from_mm100(pv)

    @prop_width.setter
    def prop_width(self, value: float | UnitT):
        """Gets/Sets the width."""
        try:
            val = value.get_value_mm100()  # type: ignore
        except AttributeError:
            val = UnitConvert.convert_mm_mm100(value)  # type: ignore
        val = max(val, 0)
        self._set(self._props.width, val)

    @property
    def prop_style(self) -> BorderLineKind:
        """Gets/Sets the style."""
        return self._prop_style

    @prop_style.setter
    def prop_style(self, value: BorderLineKind):
        """Sets the style."""
        self._set_line_style(value)
        self._prop_style = value

    @property
    def prop_transparency(self) -> Intensity:
        """Gets/Sets the transparency."""
        pv = cast(int, self._get(self._props.transparency1))
        return Intensity(pv)

    @prop_transparency.setter
    def prop_transparency(self, value: int | Intensity) -> None:
        val = Intensity(int(value))
        self._set(self._props.transparency1, val.value)
        if self._props.transparency2:
            self._set(self._props.transparency2, val.value)

    @property
    def _props(self) -> _LinePropertiesProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = _LinePropertiesProps(
                color1="BorderColor",
                color2="LineColor",
                width="BorderWidth",
                transparency1="BorderTransparency",
                transparency2="LineTransparence",
            )
        return self._props_internal_attributes

    # endregion Properties
