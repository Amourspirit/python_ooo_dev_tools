from __future__ import annotations
from typing import Any, TYPE_CHECKING


if TYPE_CHECKING:
    try:
        from typing import Protocol
    except ImportError:
        from typing_extensions import Protocol
    from ooodev.format.proto.style_t import StyleT
    from ooo.dyn.chart.chart_axis_label_position import ChartAxisLabelPosition
else:
    Protocol = object
    StyleT = Any
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
