from __future__ import annotations
from ooodev.theme.theme import ThemeBase


class ThemeSql(ThemeBase):
    """
    Theme SQL Properties.

    The properties are populated from LibreOffice theme colors.

    Automatic color values are returned with a value of ``-1``.
    All other values are positive numbers.
    """

    # region Properties
    @property
    def comment_color(self) -> int:
        """Comment color."""
        try:
            return self._comment_color
        except AttributeError:
            self._comment_color = self._get_color("SQLComment")
        return self._comment_color

    @property
    def identifier_color(self) -> int:
        """Identifier color."""
        try:
            return self._identifier_color
        except AttributeError:
            self._identifier_color = self._get_color("SQLIdentifier")
        return self._identifier_color

    @property
    def keyword_color(self) -> int:
        """Keyword color."""
        try:
            return self._keyword_color
        except AttributeError:
            self._keyword_color = self._get_color("SQLKeyword")
        return self._keyword_color

    @property
    def number_color(self) -> int:
        """Keyword color."""
        try:
            return self._number_color
        except AttributeError:
            self._number_color = self._get_color("SQLNumber")
        return self._number_color

    @property
    def operator_color(self) -> int:
        """Operator color."""
        try:
            return self._operator_color
        except AttributeError:
            self._operator_color = self._get_color("SQLOperator")
        return self._operator_color

    @property
    def parameter_color(self) -> int:
        """Parameter color."""
        try:
            return self._parameter_color
        except AttributeError:
            self._parameter_color = self._get_color("SQLParameter")
        return self._parameter_color

    @property
    def string_color(self) -> int:
        """String color."""
        try:
            return self._string_color
        except AttributeError:
            self._string_color = self._get_color("SQLString")
        return self._string_color

    # endregion Properties
