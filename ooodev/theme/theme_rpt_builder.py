from __future__ import annotations
from ooodev.theme.theme import ThemeBase
from ooodev.utils.info import Info


class ThemeRptBuilder(ThemeBase):
    """
    Theme Report Builder properties

    The properties are populated from LibreOffice theme colors.

    Automatic color values are returned with a value of ``-1``.
    All other values are positive numbers.
    """

    def _get_color(self, prop_name: str) -> int:
        val = Info.get_config(
            node_str="Color",
            node_path=f"org.openoffice.Office.ExtendedColorScheme/ExtendedColorScheme/ColorSchemes/{self._theme_name}/SunReportBuilder/Entries/{prop_name}",
        )
        return -1 if val is None else int(val)

    # region Properties
    @property
    def detail_color(self) -> int:
        """Detail color."""
        try:
            return self._detail_color
        except AttributeError:
            self._detail_color = self._get_color("Detail")
        return self._detail_color

    @property
    def page_header_color(self) -> int:
        """Page Header color."""
        try:
            return self._page_header_color
        except AttributeError:
            self._page_header_color = self._get_color("PageHeader")
        return self._page_header_color

    @property
    def page_footer_color(self) -> int:
        """Page Footer color."""
        try:
            return self._page_footer_color
        except AttributeError:
            self._page_footer_color = self._get_color("PageFooter")
        return self._page_footer_color

    @property
    def group_header_color(self) -> int:
        """Group Header color."""
        try:
            return self._group_header_color
        except AttributeError:
            self._group_header_color = self._get_color("GroupHeader")
        return self._group_header_color

    @property
    def group_footer_color(self) -> int:
        """Group Footer color."""
        try:
            return self._group_footer_color
        except AttributeError:
            self._group_footer_color = self._get_color("GroupFooter")
        return self._group_footer_color

    @property
    def column_header_color(self) -> int:
        """Column Header color."""
        try:
            return self._column_header_color
        except AttributeError:
            self._column_header_color = self._get_color("ColumnHeader")
        return self._column_header_color

    @property
    def column_footer_color(self) -> int:
        """Column Footer color."""
        try:
            return self._column_footer_color
        except AttributeError:
            self._column_footer_color = self._get_color("ColumnFooter")
        return self._column_footer_color

    @property
    def report_header_color(self) -> int:
        """Report Header color."""
        try:
            return self._report_header_color
        except AttributeError:
            self._report_header_color = self._get_color("ReportHeader")
        return self._report_header_color

    @property
    def report_footer_color(self) -> int:
        """Report Footer color."""
        try:
            return self._report_footer_color
        except AttributeError:
            self._report_footer_color = self._get_color("ReportFooter")
        return self._report_footer_color

    @property
    def overlap_control_color(self) -> int:
        """Overlapped control color."""
        try:
            return self._overlap_control_color
        except AttributeError:
            self._overlap_control_color = self._get_color("OverlappedControl")
        return self._overlap_control_color

    @property
    def text_box_bound_content_color(self) -> int:
        """Text Box Bound Content color."""
        try:
            return self._text_box_bound_content_color
        except AttributeError:
            self._text_box_bound_content_color = self._get_color("TextBoxBoundContent")
        return self._text_box_bound_content_color

    # endregion Properties
