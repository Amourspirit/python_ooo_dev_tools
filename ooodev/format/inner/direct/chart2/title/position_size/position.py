from __future__ import annotations
from typing import Any
import uno

from ooodev.format.inner.direct.chart2.position_size.position import Position as ChartShapePosition
from ooodev.units.unit_obj import UnitT


class Position(ChartShapePosition):
    """
    Positions a Title.

    .. seealso::

        - :ref:`help_chart2_format_direct_title_position_size`

    .. versionadded:: 0.9.4
    """

    def __init__(
        self,
        pos_x: float | UnitT,
        pos_y: float | UnitT,
    ) -> None:
        """
        Constructor

        Args:
            pos_x (float, UnitT): Specifies the x-coordinate of the position of the shape (in ``mm`` units) or :ref:`proto_unit_obj`.
            pos_y (float, UnitT): Specifies the y-coordinate of the position of the shape (in ``mm`` units) or :ref:`proto_unit_obj`.

        Returns:
            None:

        See Also:
            - :ref:`help_chart2_format_direct_title_position_size`
        """
        super().__init__(pos_x=pos_x, pos_y=pos_y)

    # region Overridden Methods
    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies tab properties to ``obj``

        Args:
            obj (object): UNO object.

        Returns:
            None:
        """
        props = kwargs.pop("override_dv", {})
        props.update({"AutomaticPosition": False})
        super().apply(obj=obj, override_dv=props)

    # endregion Overridden Methods
