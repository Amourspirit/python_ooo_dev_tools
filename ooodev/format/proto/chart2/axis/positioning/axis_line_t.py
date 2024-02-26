from __future__ import annotations
from typing import Any, TYPE_CHECKING

from ooodev.mock.mock_g import DOCS_BUILDING
from ooodev.format.proto.style_t import StyleT

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Protocol
    from ooo.dyn.chart.chart_axis_position import ChartAxisPosition
else:
    Protocol = object
    ChartAxisPosition = Any


class AxisLineT(StyleT, Protocol):
    """Axis Line Protocol"""

    def __init__(self, cross: ChartAxisPosition | None = ..., value: Any = ...) -> None:
        """
        Constructor

        Args:
            cross(ChartAxisPosition, optional): The position where the axis crosses the other axis.
            value (float, None, optional): The value where the axis crosses the other axis.

        Returns:
            None:
        """

        ...

    @property
    def prop_cross(self) -> ChartAxisPosition | None: ...

    @prop_cross.setter
    def prop_cross(self, value: ChartAxisPosition | None) -> None: ...

    @property
    def prop_value(self) -> Any: ...

    @prop_value.setter
    def prop_value(self, value: Any) -> None: ...
