from __future__ import annotations
from ooodev.theme.theme import ThemeBase


class ThemeCalc(ThemeBase):
    """
    Theme Calc Properties.

    The properties are populated from LibreOffice theme colors.

    Automatic color values are returned with a value of ``-1``.
    All other values are positive numbers.
    """

    # region Properties
    @property
    def detective_color(self) -> int:
        """Detective color."""
        try:
            return self._detective_color
        except AttributeError:
            self._detective_color = self._get_color("CalcDetective")
        return self._detective_color

    @property
    def detective_error_color(self) -> int:
        """Detective Error color."""
        try:
            return self._detective_error_color
        except AttributeError:
            self._detective_error_color = self._get_color("CalcDetectiveError")
        return self._detective_error_color

    @property
    def formula_color(self) -> int:
        """Formula color."""
        try:
            return self._formula_color
        except AttributeError:
            self._formula_color = self._get_color("CalcFormula")
        return self._formula_color

    @property
    def grid_color(self) -> int:
        """Grid color."""
        try:
            return self._grid_color
        except AttributeError:
            self._grid_color = self._get_color("CalcGrid")
        return self._grid_color

    @property
    def hidden_col_row_color(self) -> int:
        """Hidden column/row color."""
        try:
            return self._hidden_col_row_color
        except AttributeError:
            self._hidden_col_row_color = self._get_color("CalcHiddenColRow")
        return self._hidden_col_row_color

    @property
    def hidden_col_row_visible(self) -> bool:
        """Hidden column/row visible."""
        try:
            return self._hidden_col_row_visible
        except AttributeError:
            self._hidden_col_row_visible = self._get_visible("CalcHiddenColRow")
        return self._hidden_col_row_visible

    @property
    def notes_background_color(self) -> int:
        """Notes background color."""
        try:
            return self._notes_background_color
        except AttributeError:
            self._notes_background_color = self._get_color("CalcNotesBackground")
        return self._notes_background_color

    @property
    def page_break_auto_color(self) -> int:
        """Page Break Automatic color."""
        try:
            return self._page_break_auto_color
        except AttributeError:
            self._page_break_auto_color = self._get_color("CalcPageBreakAutomatic")
        return self._page_break_auto_color

    @property
    def page_break_manual_color(self) -> int:
        """Page Break Manual color."""
        try:
            return self._page_break_manual_color
        except AttributeError:
            self._page_break_manual_color = self._get_color("CalcPageBreakManual")
        return self._page_break_manual_color

    @property
    def protected_background_color(self) -> int:
        """Calc Protected Background color."""
        try:
            return self._protected_background_color
        except AttributeError:
            self._protected_background_color = self._get_color("CalcProtectedBackground")
        return self._protected_background_color

    @property
    def reference_color(self) -> int:
        """Reference color."""
        try:
            return self._reference_color
        except AttributeError:
            self._reference_color = self._get_color("CalcReference")
        return self._reference_color

    @property
    def text_color(self) -> int:
        """Text color."""
        try:
            return self._text_color
        except AttributeError:
            self._text_color = self._get_color("CalcText")
        return self._text_color

    @property
    def value_color(self) -> int:
        """Value color."""
        try:
            return self._value_color
        except AttributeError:
            self._value_color = self._get_color("CalcValue")
        return self._value_color

    # endregion Properties
