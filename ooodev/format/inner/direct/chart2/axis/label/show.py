from __future__ import annotations
from typing import cast, Tuple
import uno

from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase


class Show(StyleBase):
    """
    Chart Axis Label visibility.

    .. versionadded:: 0.9.4
    """

    def __init__(self, visible: bool = True) -> None:
        """
        Constructor

        Args:
            visible (bool, optional): Specifies if the Axis label is visible. Defaults to ``True``.

        Returns:
            None:
        """
        super().__init__()
        self.prop_visible = visible

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
    def prop_visible(self) -> bool:
        """Gets or Sets if the Axis Label is visible."""
        return cast(bool, self._get("DisplayLabels"))

    @prop_visible.setter
    def prop_visible(self, value: bool) -> None:
        self._set("DisplayLabels", value)

    # endregion Properties
