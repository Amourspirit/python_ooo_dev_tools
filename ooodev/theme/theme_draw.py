from __future__ import annotations
from ooodev.theme.theme import ThemeBase


class ThemeDraw(ThemeBase):
    """
    Theme Draw Properties.

    The properties are populated from LibreOffice theme colors.

    Automatic color values are returned with a value of ``-1``.
    All other values are positive numbers.
    """

    # region Properties
    @property
    def grid_color(self) -> int:
        """Grid color."""
        try:
            return self._grid_color
        except AttributeError:
            self._grid_color = self._get_color("DrawGrid")
        return self._grid_color

    @property
    def grid_visible(self) -> bool:
        """Grid visible."""
        try:
            return self._grid_visible
        except AttributeError:
            self._grid_visible = self._get_visible("DrawGrid")
        return self._grid_visible

    # endregion Properties
