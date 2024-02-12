from __future__ import annotations
from typing import Any, Tuple, cast
from enum import IntFlag
import uno

# com.sun.star.chart2.TickmarkStyle
from ooo.dyn.chart2.tickmark_style import TickmarkStyle

# com.sun.star.chart.ChartAxisMarkPosition
from ooo.dyn.chart.chart_axis_mark_position import ChartAxisMarkPosition

from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase


class MarkKind(IntFlag):
    """Tick mark Style"""

    NONE = 0
    """No marks"""
    OUTSIDE = TickmarkStyle.OUTER
    """Marks outside the axis"""
    INSIDE = TickmarkStyle.INNER
    """Marks inside the axis"""
    BOTH = TickmarkStyle.OUTER | TickmarkStyle.INNER
    """Marks inside and outside the axis"""


class IntervalMarks(StyleBase):
    """
    Chart Axis Interval Marks.

    .. seealso::

        - :ref:`help_chart2_format_direct_axis_positioning`

    .. versionadded:: 0.9.4
    """

    def __init__(
        self, major: MarkKind | None = None, minor: MarkKind | None = None, pos: ChartAxisMarkPosition | None = None
    ) -> None:
        """
        Constructor

        Args:
            major (MarkKind, optional): Specifies the major tickmark style.
            minor (MarkKind, optional): Specifies the minor tickmark style.
            pos (ChartAxisMarkPosition, optional): Specifies where to place the marks: at labels, at axis, or at axis and labels.

        Returns:
            None:

        See Also:
            - :ref:`help_chart2_format_direct_axis_positioning`
        """
        # pos, only valid when Label Position ( set using LabelPosition class) is Outside end or Outside start
        super().__init__()
        if major is not None:
            self.prop_major = major
        if minor is not None:
            self.prop_minor = minor
        if pos is not None:
            self.prop_pos = pos

    # region overrides
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.chart2.Axis",)
        return self._supported_services_values

    # endregion overrides

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
    def prop_major(self) -> MarkKind | None:
        pv = cast(int, self._get("MajorTickmarks"))
        return None if pv is None else MarkKind(pv)

    @prop_major.setter
    def prop_major(self, value: MarkKind | None) -> None:
        if value is None:
            self._remove("MajorTickmarks")
            return
        self._set("MajorTickmarks", int(value))

    @property
    def prop_minor(self) -> MarkKind | None:
        pv = cast(int, self._get("MinorTickmarks"))
        return None if pv is None else MarkKind(pv)

    @prop_minor.setter
    def prop_minor(self, value: MarkKind | None) -> None:
        if value is None:
            self._remove("MinorTickmarks")
            return
        self._set("MinorTickmarks", int(value))

    @property
    def prop_pos(self) -> ChartAxisMarkPosition | None:
        return self._get("MarkPosition")

    @prop_pos.setter
    def prop_pos(self, value: ChartAxisMarkPosition | None) -> None:
        if value is None:
            self._remove("MarkPosition")
            return
        self._set("MarkPosition", value)

    # endregion Properties
