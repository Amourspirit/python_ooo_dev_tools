from __future__ import annotations
from ooodev.theme.theme import ThemeBase


class ThemeTextDoc(ThemeBase):
    """
    Theme Text Doc (Writer) Properties.

    The properties are populated from LibreOffice theme colors.

    Automatic color values are returned with a value of ``-1``.
    All other values are positive numbers.
    """

    # region Properties
    @property
    def boundaries_color(self) -> int:
        """Boundaries color."""
        try:
            return self._boundaries_color
        except AttributeError:
            self._boundaries_color = self._get_color("DocBoundaries")
        return self._boundaries_color

    @property
    def boundaries_visible(self) -> bool:
        """Boundaries visible."""
        try:
            return self._boundaries_visible
        except AttributeError:
            self._boundaries_visible = self._get_visible("DocBoundaries")
        return self._boundaries_visible

    @property
    def doc_color(self) -> int:
        """Document color."""
        try:
            return self._doc_color
        except AttributeError:
            self._doc_color = self._get_color("DocColor")
        return self._doc_color

    @property
    def grammar_color(self) -> int:
        """Grammar mistakes color."""
        try:
            return self._grammar_color
        except AttributeError:
            self._grammar_color = self._get_color("Grammar")
        return self._grammar_color

    @property
    def direct_cursor_color(self) -> int:
        """Direct Cursor color."""
        try:
            return self._direct_cursor_color
        except AttributeError:
            self._direct_cursor_color = self._get_color("WriterDirectCursor")
        return self._direct_cursor_color

    @property
    def direct_cursor_visible(self) -> bool:
        """Direct Cursor visible."""
        try:
            return self._direct_cursor_visible
        except AttributeError:
            self._direct_cursor_visible = self._get_visible("WriterDirectCursor")
        return self._direct_cursor_visible

    @property
    def field_shadings_color(self) -> int:
        """Field Shadings color."""
        try:
            return self._field_shadings_color
        except AttributeError:
            self._field_shadings_color = self._get_color("WriterFieldShadings")
        return self._field_shadings_color

    @property
    def header_footer_mark_color(self) -> int:
        """Header and footer delimiter color."""
        try:
            return self._header_footer_mark_color
        except AttributeError:
            self._header_footer_mark_color = self._get_color("WriterHeaderFooterMark")
        return self._header_footer_mark_color

    @property
    def index_table_shadings_color(self) -> int:
        """Index and Table Shadings color."""
        try:
            return self._index_table_shadings_color
        except AttributeError:
            self._index_table_shadings_color = self._get_color("WriterIdxShadings")
        return self._index_table_shadings_color

    @property
    def index_table_shadings_visible(self) -> bool:
        """Index and Table Shadings visible."""
        try:
            return self._index_table_shadings_visible
        except AttributeError:
            self._index_table_shadings_visible = self._get_visible("WriterIdxShadings")
        return self._index_table_shadings_visible

    @property
    def page_columns_breaks_color(self) -> int:
        """Page and Columns breaks color."""
        try:
            return self._page_columns_breaks_color
        except AttributeError:
            self._page_columns_breaks_color = self._get_color("WriterPageBreaks")
        return self._page_columns_breaks_color

    @property
    def script_indicator_color(self) -> int:
        """Script Indicator color."""
        try:
            return self._script_indicator_color
        except AttributeError:
            self._script_indicator_color = self._get_color("WriterScriptIndicator")
        return self._script_indicator_color

    @property
    def section_boundaries_color(self) -> int:
        """Section Boundaries color."""
        try:
            return self._section_boundaries_color
        except AttributeError:
            self._section_boundaries_color = self._get_color("WriterSectionBoundaries")
        return self._section_boundaries_color

    @property
    def section_boundaries_visible(self) -> bool:
        """Direct Cursor visible."""
        try:
            return self._section_boundaries_visible
        except AttributeError:
            self._section_boundaries_visible = self._get_visible("WriterSectionBoundaries")
        return self._section_boundaries_visible

    @property
    def grid_color(self) -> int:
        """Grid color."""
        try:
            return self._grid_color
        except AttributeError:
            self._grid_color = self._get_color("WriterTextGrid")
        return self._grid_color

    # endregion Properties
