from __future__ import annotations
from typing import Any, TYPE_CHECKING


if TYPE_CHECKING:
    try:
        from typing import Protocol
    except ImportError:
        from typing_extensions import Protocol
    from ooodev.format.proto.style_t import StyleT
    from ooo.dyn.chart.chart_axis_position import ChartAxisPosition
else:
    Protocol = object
    StyleT = Any
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
