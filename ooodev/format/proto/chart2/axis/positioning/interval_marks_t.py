from __future__ import annotations
from typing import Any, TYPE_CHECKING

from ooodev.mock.mock_g import DOCS_BUILDING
from ooodev.format.proto.style_t import StyleT

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Protocol
    from ooodev.format.inner.direct.chart2.axis.positioning.interval_marks import MarkKind
    from ooo.dyn.chart.chart_axis_mark_position import ChartAxisMarkPosition
else:
    Protocol = object
    MarkKind = Any
    ChartAxisMarkPosition = Any


class IntervalMarksT(StyleT, Protocol):
    """Axis Interval Marks Protocol"""

    def __init__(
        self, major: MarkKind | None = ..., minor: MarkKind | None = ..., pos: ChartAxisMarkPosition | None = ...
    ) -> None:
        """
        Constructor

        Args:
            major (MarkKind, optional): Specifies the major tickmark style.
            minor (MarkKind, optional): Specifies the minor tickmark style.
            pos (ChartAxisMarkPosition, optional): Specifies where to place the marks: at labels, at axis, or at axis and labels.

        Returns:
            None:
        """

        ...

    @property
    def prop_major(self) -> MarkKind | None: ...

    @prop_major.setter
    def prop_major(self, value: MarkKind | None) -> None: ...

    @property
    def prop_minor(self) -> MarkKind | None: ...

    @prop_minor.setter
    def prop_minor(self, value: MarkKind | None) -> None: ...

    @property
    def prop_pos(self) -> ChartAxisMarkPosition | None: ...

    @prop_pos.setter
    def prop_pos(self, value: ChartAxisMarkPosition | None) -> None: ...
