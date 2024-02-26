from __future__ import annotations
from ooodev.theme.theme import ThemeBase


class ThemeHtml(ThemeBase):
    """
    Theme HTML properties.

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
            self._comment_color = self._get_color("HTMLComment")
        return self._comment_color

    @property
    def keyword_color(self) -> int:
        """Keyword color."""
        try:
            return self._keyword_color
        except AttributeError:
            self._keyword_color = self._get_color("HTMLKeyword")
        return self._keyword_color

    @property
    def sgml_color(self) -> int:
        """SGML syntax highlighting color."""
        try:
            return self._sgml_color
        except AttributeError:
            self._sgml_color = self._get_color("HTMLSGML")
        return self._sgml_color

    @property
    def unknown_color(self) -> int:
        """Unknown color."""
        try:
            return self._unknown_color
        except AttributeError:
            self._unknown_color = self._get_color("HTMLUnknown")
        return self._unknown_color

    # endregion Properties
