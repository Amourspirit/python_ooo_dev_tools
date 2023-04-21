from __future__ import annotations
from typing import Any
import uno

from ooodev.events.args.key_val_cancel_args import KeyValCancelArgs

from ...position_size.position import Position as ChartShapePosition


class Position(ChartShapePosition):
    """
    Positions a Title.

    .. versionadded:: 0.9.4
    """

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
