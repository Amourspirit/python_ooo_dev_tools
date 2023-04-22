from __future__ import annotations
from typing import cast, Tuple
import uno

from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase


class TextFlow(StyleBase):
    """
    Chart Axis Label Text Flow.

    .. versionadded:: 0.9.4
    """

    def __init__(self, overlap: bool | None = None, brk: bool | None = None) -> None:
        """
        Constructor

        Args:
            overlap (bool, optional): Specifies that the text in cells may overlap other cells. This can be especially useful if there is a lack of space. This option is not available with different title directions.
            break (bool, optional): Specifies if a text break is allowed.

        Returns:
            None:
        """
        super().__init__()
        if overlap is not None:
            self.prop_overlap = overlap
        if brk is not None:
            self.prop_brk = brk

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
    def prop_overlap(self) -> bool | None:
        """Gets or Sets if the text in cells may overlap other cells."""
        return cast(bool, self._get("TextOverlap"))

    @prop_overlap.setter
    def prop_overlap(self, value: bool | None) -> None:
        self._set("TextOverlap", value)

    @property
    def prop_brk(self) -> bool | None:
        """Gets or Sets if a text break is allowed."""
        return cast(bool, self._get("TextBreak"))

    @prop_brk.setter
    def prop_brk(self, value: bool | None) -> None:
        self._set("TextBreak", value)

    # endregion Properties
