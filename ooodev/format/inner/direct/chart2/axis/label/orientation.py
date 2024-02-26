from __future__ import annotations
from typing import cast, Tuple
import uno

from ooodev.format.inner.direct.chart2.title.alignment.direction import DirectionModeKind
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.units.angle import Angle


class Orientation(StyleBase):
    """
    Chart Axis Text orientation.

    .. versionadded:: 0.9.4
    """

    def __init__(
        self, angle: int | Angle | None = None, mode: DirectionModeKind | None = None, vertical: bool | None = None
    ) -> None:
        """
        Constructor

        Args:
            angle (int, Angle, optional): Rotation in degrees of the text.
            mode (DirectionModeKind, optional): Specifies the writing direction.
            vertical (bool, optional): Specifies if the text is vertically stacked.

        Returns:
            None:

        Note:
            When ``vertical`` is ``True`` the ``angle`` is ignored.
        """
        super().__init__()
        if angle is not None:
            self.prop_angle = angle
        if mode is not None:
            self.prop_mode = mode
        if vertical is not None:
            self.prop_vertical = vertical

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
    def prop_vertical(self) -> bool | None:
        """Gets/Sets if the text is vertically stacked."""
        pv = cast(bool, self._get("StackCharacters"))
        return None if pv is None else pv

    @prop_vertical.setter
    def prop_vertical(self, value: bool | None) -> None:
        if value is None:
            self._remove("StackCharacters")
            return
        self._set("StackCharacters", value)

    # endregion Properties
