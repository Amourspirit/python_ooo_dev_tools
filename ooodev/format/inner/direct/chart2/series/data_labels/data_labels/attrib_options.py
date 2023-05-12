from __future__ import annotations
from enum import Enum
from typing import Tuple, cast
import uno
from ooo.dyn.chart.data_label_placement import DataLabelPlacement
from ooodev.format.inner.kind.format_kind import FormatKind

from ooodev.format.inner.style_base import StyleBase


class PlacementKind(Enum):
    """Placement Kind."""

    ABOVE = DataLabelPlacement.TOP
    BELOW = DataLabelPlacement.BOTTOM
    CENTER = DataLabelPlacement.CENTER
    OUTSIDE = DataLabelPlacement.OUTSIDE
    INSIDE = DataLabelPlacement.INSIDE
    NEAR_ORIGIN = DataLabelPlacement.NEAR_ORIGIN

    def __int__(self) -> int:
        return self.value


class SeparatorKind(Enum):
    """Separator Kind."""

    COMMA = ", "
    SPACE = " "
    SEMICOLON = "; "
    NEW_LINE = "\n"
    PERIOD = ". "

    def __str__(self) -> str:
        return self.value


class AttribOptions(StyleBase):
    """
    Chart Data Series, Data Labels Text Attribute Options.

    .. seealso::

        - :ref:`help_chart2_format_direct_series_labels_data_labels`
    """

    def __init__(self, placement: PlacementKind | None = None, separator: SeparatorKind | None = None) -> None:
        """
        Constructor

        Args:
            placement (PlacementKind, optional): Specifies the placement of data labels relative to the objects.
            separator (SeparatorKind, optional): Specifies the separator between multiple text strings for the same object.

        Returns:
            None:

        See Also:
            - :ref:`help_chart2_format_direct_series_labels_data_labels`
        """
        super().__init__()
        if placement is not None:
            self.prop_placement = placement
        if separator is not None:
            self.prop_separator = separator

    # region overrides methods

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.chart2.DataSeries",
                "com.sun.star.chart2.DataPointProperties",
            )
        return self._supported_services_values

    # endregion overrides methods

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
    def prop_placement(self) -> PlacementKind | None:
        """Gets or sets the placement kind"""
        pv = cast(int, self._get("LabelPlacement"))
        return None if pv is None else PlacementKind(pv)

    @prop_placement.setter
    def prop_placement(self, value: PlacementKind | None) -> None:
        """Sets the placement kind"""
        if value is None:
            self._remove("LabelPlacement")
            return
        self._set("LabelPlacement", int(value))

    @property
    def prop_separator(self) -> SeparatorKind | None:
        """Gets or sets the separator kind"""
        pv = cast(str, self._get("LabelSeparator"))
        return None if pv is None else SeparatorKind(pv)

    @prop_separator.setter
    def prop_separator(self, value: SeparatorKind | None) -> None:
        """Sets the separator kind"""
        if value is None:
            self._remove("LabelSeparator")
            return
        self._set("LabelSeparator", str(value))

    # endregion Properties
