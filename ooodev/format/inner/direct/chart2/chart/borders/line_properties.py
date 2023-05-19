from __future__ import annotations
import uno
from typing import Tuple, cast, overload
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.preset.preset_border_line import BorderLineKind, get_preset_border_line_props
from ooodev.format.inner.style_base import StyleBase
from ooodev.units import UnitConvert, UnitMM
from ooodev.units import UnitObj
from ooodev.utils.color import Color
from ooodev.utils.data_type.intensity import Intensity


class LineProperties(StyleBase):
    """
    This class represents the line properties of a chart borders line properties.

    .. seealso::

        - :ref:`help_chart2_format_direct_general_borders`
    """

    def __init__(
        self,
        style: BorderLineKind = BorderLineKind.CONTINUOUS,
        color: Color = Color(0),
        width: float | UnitObj = 0,
        transparency: int | Intensity = 0,
    ) -> None:
        """
        Constructor.

        Args:
            style (BorderLineKind): Line style. Defaults to ``BorderLineKind.CONTINUOUS``.
            color (Color, optional): Line Color. Defaults to ``Color(0)``.
            width (float | UnitObj, optional): Line Width (in ``mm`` units) or :ref:`proto_unit_obj`. Defaults to ``0``.
            transparency (int | Intensity, optional): Line transparency from ``0`` to ``100``. Defaults to ``0``.

        Returns:
            None:

        See Also:
            - :ref:`help_chart2_format_direct_general_borders`
        """
        super().__init__()
        self._prop_style = style
        self._set_line_style(style)
        self.prop_color = color
        self.prop_width = width
        self.prop_transparency = transparency

    # region Private Methods

    def _set_line_style(self, style: BorderLineKind):
        props = get_preset_border_line_props(kind=style)
        self._set("LineCap", props.line_cap)
        self._set("LineDash", props.line_dash)
        self._set("LineDashName", props.line_dash_name)
        self._set("LineJoint", props.line_joint)
        self._set("LineStyle", props.line_style)

    # endregion Private Methods

    # region Overridden Methods
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.chart2.Legend",
                "com.sun.star.chart2.PageBackground",
                "com.sun.star.chart2.Title",
                "com.sun.star.drawing.LineProperties",
                "com.sun.star.chart2.Axis",
            )
        return self._supported_services_values

    # region copy()
    @overload
    def copy(self) -> LineProperties:
        ...

    @overload
    def copy(self, **kwargs) -> LineProperties:
        ...

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
        return self._get("LineColor")

    @prop_color.setter
    def prop_color(self, value: Color):
        value = max(value, 0)  # type: ignore
        self._set("LineColor", value)

    @property
    def prop_width(self) -> UnitMM:
        pv = cast(int, self._get("LineWidth"))
        return UnitMM.from_mm100(pv)

    @prop_width.setter
    def prop_width(self, value: float | UnitObj):
        """Gets/Sets the width."""
        try:
            val = value.get_value_mm100()  # type: ignore
        except AttributeError:
            val = UnitConvert.convert_mm_mm100(value)  # type: ignore
        val = max(val, 0)
        self._set("LineWidth", val)

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
        pv = cast(int, self._get("LineTransparence"))
        return Intensity(pv)

    @prop_transparency.setter
    def prop_transparency(self, value: int | Intensity) -> None:
        val = Intensity(int(value))
        self._set("LineTransparence", val.value)

    # endregion Properties
