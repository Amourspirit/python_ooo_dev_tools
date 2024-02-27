from __future__ import annotations
from typing import Any, TYPE_CHECKING

from ooodev.mock.mock_g import DOCS_BUILDING
from ooodev.format.proto.style_t import StyleT

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Protocol
    from ooo.dyn.chart.chart_axis_label_position import ChartAxisLabelPosition
else:
    Protocol = object
    ChartAxisLabelPosition = Any


class LabelPositionT(StyleT, Protocol):
    """Axis Label Position Protocol"""

    def __init__(self, pos: ChartAxisLabelPosition = ...) -> None:
        """
        Constructor

        Args:
            pos (ChartAxisLabelPosition, optional): Specifies where to place the labels: near axis, near axis (other side), outside start, or outside end.
                Default is ``ChartAxisLabelPosition.NEAR_AXIS``.

        Returns:
            None:
        """

        ...

    @property
    def prop_pos(self) -> ChartAxisLabelPosition: ...

    @prop_pos.setter
    def prop_pos(self, value: ChartAxisLabelPosition) -> None: ...
