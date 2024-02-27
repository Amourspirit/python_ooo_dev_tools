from __future__ import annotations
import contextlib
from typing import Any, cast, Tuple
import uno
from ooodev.events.args.key_val_cancel_args import KeyValCancelArgs

from ooodev.format.inner.direct.chart2.title.alignment.direction import DirectionModeKind
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.units.angle import Angle


class Orientation(StyleBase):
    """
    Series data points Text orientation.

    .. seealso::

        - :ref:`help_chart2_format_direct_series_labels_data_labels`

    .. versionadded:: 0.9.4
    """

    def __init__(
        self, angle: int | Angle | None = None, mode: DirectionModeKind | None = None, leaders: bool | None = None
    ) -> None:
        """
        Constructor

        Args:
            angle (int, Angle, optional): Rotation in degrees of the text.
            mode (DirectionModeKind, optional): Specifies the writing direction.
            leaders (bool, optional): Leader Lines. Connect displaced data points to data points.

        Returns:
            None:

        Note:
            When setting a data point ``leaders`` is ignored. To set ``leaders`` set it on the data series.

        See Also:
            - :ref:`help_chart2_format_direct_series_labels_data_labels`
        """
        super().__init__()
        if angle is not None:
            self.prop_angle = angle
        if mode is not None:
            self.prop_mode = mode
        if leaders is not None:
            self.prop_leaders = leaders

    # region overrides
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.chart2.DataSeries",
                "com.sun.star.chart2.DataPointProperties",
            )
        return self._supported_services_values

    # endregion overrides

    # region on events
    def on_property_set_error(self, source: Any, event_args: KeyValCancelArgs) -> None:
        if event_args.key == "ShowCustomLeaderLines":
            # ShowCustomLeaderLines must be set on data series and not data point
            with contextlib.suppress(Exception):
                if event_args.event_data.getImplementationName() == "com.sun.star.comp.chart.DataPoint":
                    event_args.handled = True
        super().on_property_set_error(source, event_args)

    # endregion on events

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
        """Gets/Sets the rotation of the text."""
        # angle is stored as float but the dialog only allows integer values
        pv = cast(float, self._get("TextRotation"))
        return None if pv is None else Angle(int(pv))

    @prop_angle.setter
    def prop_angle(self, value: int | Angle | None) -> None:
        if value is None:
            self._remove("TextRotation")
            return
        val = Angle(int(value))
        # angle is stored as float but the dialog only allows integer values
        self._set("TextRotation", float(val.value))

    @property
    def prop_mode(self) -> DirectionModeKind | None:
        """Gets/Sets writing mode."""
        pv = cast(int, self._get("WritingMode"))
        return None if pv is None else DirectionModeKind(pv)

    @prop_mode.setter
    def prop_mode(self, value: DirectionModeKind | None):
        if value is None:
            self._remove("WritingMode")
            return
        self._set("WritingMode", value.value)

    @property
    def prop_leaders(self) -> bool | None:
        """Gets/Sets whether to show leaders."""
        return cast(bool, self._get("ShowCustomLeaderLines"))

    @prop_leaders.setter
    def prop_leaders(self, value: bool | None) -> None:
        if value is None:
            self._remove("ShowCustomLeaderLines")
            return
        self._set("ShowCustomLeaderLines", value)

    # endregion Properties
