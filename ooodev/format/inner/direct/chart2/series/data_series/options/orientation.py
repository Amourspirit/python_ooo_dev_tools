from __future__ import annotations
import contextlib
from typing import Any, Tuple, TypeVar, cast, overload
import uno
from com.sun.star.chart2 import XChartDocument
from com.sun.star.chart import XAxisYSupplier

from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.loader import lo as mLo
from ooodev.units.angle import Angle

_TOrientation = TypeVar("_TOrientation", bound="Orientation")


class Orientation(StyleBase):
    """
    Chart Orientation.

    Available for pie and donut charts.

    .. seealso::

        - :ref:`help_chart2_format_direct_series_series_options`
    """

    def __init__(
        self,
        chart_doc: XChartDocument,
        clockwise: bool | None = None,
        angle: int | Angle | None = None,
    ) -> None:
        """
        Constructor

        Args:
            chart_doc (XChartDocument): Chart document.
            clockwise (bool, optional): Specifies he default direction in which the pieces of a pie chart are ordered is counterclockwise.
                Set to ``True`` to enable the Clockwise direction to draw the pieces in opposite direction.
            angle (int, Angle, optional): Sets the starting angle of a pie or donut chart.

        Returns:
            None:

        See Also:
            - :ref:`help_chart2_format_direct_series_series_options`
        """
        # clockwise default is false
        self._chart_doc = chart_doc
        super().__init__()
        if clockwise is not None:
            self.prop_clockwise = clockwise
        if angle is not None:
            self.prop_angle = angle

    # region Overrides
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.chart2.DataSeries",)
        return self._supported_services_values

    def _is_valid_obj(self, obj: Any) -> bool:
        d_types = ("com.sun.star.chart.PieDiagram", "com.sun.star.chart.DonutDiagram")
        with contextlib.suppress(Exception):
            dia = self._chart_doc.getDiagram()  # type: ignore
            d_type = dia.DiagramType
            return d_type in d_types
        return False

    # region apply()
    @overload
    def apply(self, obj: Any) -> None: ...

    @overload
    def apply(self, obj: Any, **kwargs) -> None: ...

    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies styles to object

        Args:
            obj (object): UNO Object that styles are to be applied.

        Returns:
            None:
        """
        if not self._is_valid_obj(obj):
            mLo.Lo.print(f"{self.__class__.__name__}. Only available for pie and donut charts.")
            return

        try:
            diagram = self._chart_doc.getDiagram()  # type: ignore
            angle = self.prop_angle
            if angle is not None:
                super().apply(obj=diagram, validate=False, override_dv={"StartingAngle": angle.value}, **kwargs)

            clockwise = self.prop_clockwise
            if clockwise is not None:
                supplier = mLo.Lo.qi(XAxisYSupplier, diagram, True)
                axis_props = supplier.getYAxis()
                super().apply(obj=axis_props, validate=False, override_dv={"ReverseDirection": clockwise}, **kwargs)
        except Exception as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply() - Unable to get chart diagram")
            mLo.Lo.print(f"  Error: {e}")
            return

    # endregion apply()

    # region Copy()
    @overload
    def copy(self: _TOrientation) -> _TOrientation: ...

    @overload
    def copy(self: _TOrientation, **kwargs) -> _TOrientation: ...

    def copy(self: _TOrientation, **kwargs) -> _TOrientation:
        """Gets a copy of instance as a new instance"""
        cp = super().copy(**kwargs)
        cp._chart_doc = self._chart_doc
        return cp

    # endregion Copy()
    # endregion Overrides

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
    def prop_angle(self) -> Angle | None:
        """Gets or sets the angle property"""
        pv = cast(int, self._get("StartingAngle"))
        return None if pv is None else Angle(pv)

    @prop_angle.setter
    def prop_angle(self, value: Angle | int | None) -> None:
        if value is None:
            self._remove("StartingAngle")
            return
        self._set("StartingAngle", Angle(int(value)).value)

    @property
    def prop_clockwise(self) -> bool | None:
        """Gets or sets the clockwise property"""
        return self._get("ReverseDirection")

    @prop_clockwise.setter
    def prop_clockwise(self, value: bool | None) -> None:
        if value is None:
            self._remove("ReverseDirection")
            return
        self._set("ReverseDirection", value)

    # endregion Properties
