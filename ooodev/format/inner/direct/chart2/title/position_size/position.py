from __future__ import annotations
from typing import Any
import uno

from ooodev.events.args.key_val_cancel_args import KeyValCancelArgs
from ooodev.format.inner.direct.chart2.position_size.position import Position as ChartShapePosition
from ooodev.units import UnitObj


class Position(ChartShapePosition):
    """
    Positions a Title.

    .. seealso::

        - :ref:`help_chart2_format_direct_title_position_size`

    .. versionadded:: 0.9.4
    """

    def __init__(
        self,
        pos_x: float | UnitObj,
        pos_y: float | UnitObj,
    ) -> None:
        """
        Constructor

        Args:
            pos_x (float, UnitObj): Specifies the x-coordinate of the position of the shape (in ``mm`` units) or :ref:`proto_unit_obj`.
            pos_y (float, UnitObj): Specifies the y-coordinate of the position of the shape (in ``mm`` units) or :ref:`proto_unit_obj`.

        Returns:
            None:

        See Also:
            - :ref:`help_chart2_format_direct_title_position_size`
        """
        super().__init__(pos_x=pos_x, pos_y=pos_y)

    # region Overridden Methods
    def apply(self, obj: object, **kwargs) -> None:
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

    def on_property_setting(self, source: Any, event_args: KeyValCancelArgs) -> None:
        super().on_property_setting(source, event_args)

    # endregion Overridden Methods
