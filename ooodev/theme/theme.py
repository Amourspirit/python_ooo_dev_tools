from __future__ import annotations
from typing import cast, NamedTuple, Dict, TYPE_CHECKING
from abc import ABC
from enum import Enum

from ooodev.loader.lo import Lo
from ooodev.utils.info import Info
from ooodev.utils.sys_info import SysInfo
from ooodev.utils import color as mColor
from ooodev.io.log.named_logger import NamedLogger
from ooodev.exceptions import ex as mEx

if TYPE_CHECKING:
    from com.sun.star.awt import XWindow
    from com.sun.star.frame import XFrame
    from com.sun.star.awt import XStyleSettings


class ThemeKind(Enum):
    """Theme Kind Lookup"""

    AUTOMATIC = "COLOR_SCHEME_LIBREOFFICE_AUTOMATIC"
    LIBRE_OFFICE = "LibreOffice"
    LIBRE_OFFICE_DARK = "LibreOffice Dark"

    def __str__(self) -> str:
        return self.value


# https://pypi.org/project/darkdetect/


# <item oor:path="/org.openoffice.Office.Common/Misc"><prop oor:name="ApplicationAppearance" oor:op="fuse"><value>1</value></prop></item>
# ApplicationAppearance = 1 Light
# ApplicationAppearance = 2 Dark
# ApplicationAppearance = 0 System
class ThemeColorKind(Enum):
    UNKNOWN = -1
    SYSTEM = 0
    LIGHT = 1
    DARK = 2


class ThemeDefault(NamedTuple):
    """Theme Defaults"""

    prop_name: str
    light_color: int
    dark_color: int
    visible: bool = True


DEFAULT_THEME_DETAILS: Dict[str, ThemeDefault] = {
    # Calc
    "CalcGrid": ThemeDefault("grid_color", 0xFFFFFF, 0x1C1C1C),
    "CalcPageBreak": ThemeDefault("page_break_color", 0xFFFFFF, 0x000000),
    "CalcPageBreakManual": ThemeDefault("page_break_manual_color", 0x2300DC, 0x2300DC),
    "CalcPageBreakAutomatic": ThemeDefault("page_break_auto_color", 0x666666, 0x666666),
    "CalcHiddenColRow": ThemeDefault("hidden_col_row_color", 0x2300DC, 0x2300DC, False),
    "CalcTextOverflow": ThemeDefault("text_overflow_color", 0xFF0000, 0xFF0000, True),
    "CalcComments": ThemeDefault("comments_color", 0xFF00FF, 0xFF00FF),
    "CalcDetective": ThemeDefault("detective_color", 0x0000FF, 0x355269),
    "CalcDetectiveError": ThemeDefault("detective_error_color", 0xFF0000, 0xC9211E),
    "CalcReference": ThemeDefault("reference_color", 0xEF0FFF, 0x0D23D5),
    "CalcNotesBackground": ThemeDefault("notes_background_color", 0xFFFFC0, 0xE8A202),
    "CalcValue": ThemeDefault("value_color", 0x0000FF, 0x729FCF),
    "CalcFormula": ThemeDefault("formula_color", 0x008000, 0x77BC65),
    "CalcText": ThemeDefault("text_color", 0x000000, 0xEEEEEE),
    "CalcProtectedBackground": ThemeDefault("protected_background_color", 0xC0C0C0, 0x1C1C1C),
    # Draw
    "DrawGrid": ThemeDefault("grid_color", 0x666666, 0x666666, True),
    # General
    "DocColor": ThemeDefault("doc_background_color", 0xFFFFFF, 0x1C1C1C),
    "DocBoundaries": ThemeDefault("doc_boundaries_color", 0xC0C0C0, 0x808080, True),
    "AppBackground": ThemeDefault("background_color", 0xFFFFFF, 0xF7F7F7),
    "ObjectBoundaries": ThemeDefault("object_boundaries_color", 0xC0C0C0, 0x808080, True),
    "TableBoundaries": ThemeDefault("table_boundaries_visible", 0xC0C0C0, 0x1C1C1C, True),
    "FontColor": ThemeDefault("font_color", 0x000000, 0x000000),
    "Links": ThemeDefault("links_color", 0x007AA6, 0x007AA6, False),
    "LinksVisited": ThemeDefault("links_visited_color", 0x002F40, 0x002F40, False),
    "Spell": ThemeDefault("spell_color", 0xFF0000, 0xC9211E),
    "SmartTags": ThemeDefault("smart_tags_color", 0xFF00FF, 0x780373),
    "Shadow": ThemeDefault("shadow_color", 0x808080, 0x1C1C1C, True),
    # HTML
    "HTMLSGML": ThemeDefault("sgml_color", 0x0000FF, 0x0000FF),
    "HTMLComment": ThemeDefault("comment_color", 0x00FF00, 0x00FF00),
    "HTMLKeyword": ThemeDefault("keyword_color", 0xFF0000, 0xFF0000),
    "HTMLUnknown": ThemeDefault("unknown_color", 0x808080, 0x808080),
    # Report Builder
    "Detail": ThemeDefault("detail_color", 0xF1C4A2, 0xF1C4A2),
    "PageFooter": ThemeDefault("page_footer_color", 0xF0C158, 0xF0C158),
    "PageHeader": ThemeDefault("page_header_color", 0xF0C158, 0xF0C158),
    "GroupFooter": ThemeDefault("group_footer_color", 0xAAC1D2, 0xAAC1D2),
    "GroupHeader": ThemeDefault("group_header_color", 0xAAC1D2, 0xAAC1D2),
    "ColumnFooter": ThemeDefault("column_footer_color", 0xAAC1D2, 0xAAC1D2),
    "ColumnHeader": ThemeDefault("column_header_color", 0xAAC1D2, 0xAAC1D2),
    "ReportFooter": ThemeDefault("report_footer_color", 0x7FA04C, 0x7FA04C),
    "ReportHeader": ThemeDefault("report_header_color", 0x7FA04C, 0x7FA04C),
    "OverlappedControl": ThemeDefault("overlap_control_color", 0xFF3366, 0xFF3366),
    "TextBoxBoundContent": ThemeDefault("text_box_bound_content_color", 0x000000, 0x000000),
    # SQL
    "SQLIdentifier": ThemeDefault("identifier_color", 0x009900, 0x009900),
    "SQLNumber": ThemeDefault("number_color", 0x000000, 0x000000),
    "SQLString": ThemeDefault("string_color", 0xCE7B00, 0xCE7B00),
    "SQLOperator": ThemeDefault("operator_color", 0x000000, 0x000000),
    "SQLKeyword": ThemeDefault("keyword_color", 0x0000E6, 0x0000E6),
    "SQLParameter": ThemeDefault("parameter_color", 0x259D9D, 0x259D9D),
    "SQLComment": ThemeDefault("comment_color", 0x808080, 0x808080),
    # Writer
    "WriterTextGrid": ThemeDefault("grid_color", 0xC0C0C0, 0x808080),
    "WriterFieldShadings": ThemeDefault("field_shadings_color", 0xC0C0C0, 0xC0C0C0, True),
    "WriterIdxShadings": ThemeDefault("index_table_shadings_color", 0xC0C0C0, 0x1C1C1C, True),
    "Grammar": ThemeDefault("grammar_color", 0x0000FF, 0x729FCF),
    "WriterScriptIndicator": ThemeDefault("script_indicator_color", 0x008000, 0x1E6A39),
    "WriterSectionBoundaries": ThemeDefault("section_boundaries_color", 0xC0C0C0, 0x666666, True),
    "WriterHeaderFooterMark": ThemeDefault("header_footer_mark_color", 0x0369A3, 0xB4C7DC),
    "WriterPageBreaks": ThemeDefault("page_columns_breaks_color", 0x000080, 0x729FCF),
    "WriterDirectCursor": ThemeDefault("direct_cursor_color", 0x000000, 0x000000, True),
}


class ThemeBase(ABC):
    def __init__(self, theme_name: ThemeKind | str = "") -> None:
        """
        Constructor

        Args:
            theme_name (ThemeKind | str, optional): Theme Name. If omitted then the current LibreOffice Theme is used.

        Returns:
            None:
        """
        self._log = NamedLogger(self.__class__.__name__)
        self._log.debug("Init")
        self._automatic_theme = False
        if theme_name == "" or theme_name == ThemeKind.AUTOMATIC:
            theme_name = Info.get_office_theme()
            self._automatic_theme = True

        if not theme_name:
            self._log.error("No theme name has been found.")
            raise ValueError("No theme name has been found,")

        self._theme_name = str(theme_name)

        if self._automatic_theme is False:
            if Info.is_office_theme(self._theme_name) is False:
                self._log.error(f"Invalid theme name: {self._theme_name}")
                raise ValueError(f"Invalid theme name: {self._theme_name}")

        self._log.debug(f"Theme Name: {theme_name}")
        self._theme_color_kind = None
        self._actual_theme_color_kind = None

    def _get_color(self, prop_name: str) -> int:
        global DEFAULT_THEME_DETAILS

        try:
            # search the ExtendedColorScheme first.
            try:
                val = Info.get_config(
                    node_str="Color",
                    node_path=f"/org.openoffice.Office.ExtendedColorScheme/ExtendedColorScheme/ColorSchemes/{self._theme_name}/{prop_name}",
                )
                self._log.debug(f"Color for {prop_name} is {val} and was found in ExtendedColorScheme configuration")
            except mEx.ConfigError:
                self._log.debug(f"Color for {prop_name} not in ExtendedColorScheme configuration")
                val = None
            if val is None:
                try:
                    val = Info.get_config(
                        node_str="Color",
                        node_path=f"/org.openoffice.Office.UI/ColorScheme/ColorSchemes/{self._theme_name}/{prop_name}",
                    )
                    self._log.debug(f"Color for {prop_name} is {val} and was found in ColorSchemes configuration")
                except mEx.ConfigError:
                    self._log.debug(f"Color for {prop_name} not in ColorSchemes configuration")
                    val = None

            if val is None:
                if prop_name in DEFAULT_THEME_DETAILS:
                    if self.actual_theme_color_kind == ThemeColorKind.DARK:
                        return DEFAULT_THEME_DETAILS[prop_name].dark_color
                    else:
                        return DEFAULT_THEME_DETAILS[prop_name].light_color
                return -1

            return int(val)
        except Exception:
            self._log.exception(f"Error getting color for {prop_name}")
            return -1

    def _get_visible(self, prop_name: str) -> bool:
        global DEFAULT_THEME_DETAILS
        try:
            # search the ExtendedColorScheme first.
            try:
                val = Info.get_config(
                    node_str="IsVisible",
                    node_path=f"/org.openoffice.Office.ExtendedColorScheme/ExtendedColorScheme/ColorSchemes/{self._theme_name}/{prop_name}",
                )
                self._log.debug(f"Color for {prop_name} is {val} and was found in ExtendedColorScheme configuration")
            except mEx.ConfigError:
                self._log.debug(f"Color for {prop_name} not in ExtendedColorScheme configuration")
                val = None
            if val is None:
                try:
                    val = Info.get_config(
                        node_str="IsVisible",
                        node_path=f"/org.openoffice.Office.UI/ColorScheme/ColorSchemes/{self._theme_name}/{prop_name}",
                    )
                    self._log.debug(f"Visible for {prop_name} is {val} and was found in configuration")
                except mEx.ConfigError:
                    self._log.debug(f"Visible for {prop_name} not in configuration")
                    val = None

            if val is None:
                if prop_name in DEFAULT_THEME_DETAILS:
                    return DEFAULT_THEME_DETAILS[prop_name].visible
                return False

            return bool(val)
        except Exception:
            return False

    def get_actual_theme_color_kind(self) -> ThemeColorKind:
        """
        Get Theme Color Kind from configuration or current Document.

        On some version of Linux the color is not correct when LibreOffice is set to use the system color scheme.
        This method will return the actual color scheme being used.

        Returns:
            ThemeColorKind: Theme Color Kind
        """
        if self._actual_theme_color_kind is not None:
            return self._actual_theme_color_kind
        try:
            current_kind = self.get_theme_color_kind()
            if current_kind == ThemeColorKind.DARK or current_kind == ThemeColorKind.LIGHT:
                self._actual_theme_color_kind = current_kind
                return self._actual_theme_color_kind

            if SysInfo.get_platform() != SysInfo.PlatformEnum.LINUX:
                self._actual_theme_color_kind = current_kind
            else:
                if self.is_document_dark():
                    self._actual_theme_color_kind = ThemeColorKind.DARK
                else:
                    self._actual_theme_color_kind = ThemeColorKind.LIGHT
        except Exception:
            self._actual_theme_color_kind = ThemeColorKind.UNKNOWN
        return self._actual_theme_color_kind

    def get_theme_color_kind(self) -> ThemeColorKind:
        """
        Get Theme Color Kind from configuration.

        Returns:
            ThemeColorKind: Theme Color Kind
        """
        if self._theme_color_kind is not None:
            return self._theme_color_kind
        try:
            val = Info.get_config(
                node_str="ApplicationAppearance",
                node_path="org.openoffice.Office.Common/Misc",
            )
            if val == 1:
                self._theme_color_kind = ThemeColorKind.LIGHT
            elif val == 2:
                self._theme_color_kind = ThemeColorKind.DARK
            elif val == 0:
                self._theme_color_kind = ThemeColorKind.SYSTEM
            else:
                self._theme_color_kind = ThemeColorKind.UNKNOWN
        except Exception:
            self._theme_color_kind = ThemeColorKind.UNKNOWN
        return self._theme_color_kind

    def is_document_dark(self) -> bool:
        """
        Is Document set to a dark color scheme.

        This method attempts to get the ``LightColor`` from the current document and determine if the document is dark.

        Returns:
            bool: ``True`` if document is dark, ``False`` otherwise.
        """
        try:
            doc = Lo.current_doc
            controller = cast("XFrame", doc.get_current_controller())
            window = cast("XWindow", controller.ComponentWindow)  # type: ignore
            style_settings = cast("XStyleSettings", window.StyleSettings)  # type: ignore
            rgb_color = mColor.RGB.from_color(style_settings.LightColor)  # type: ignore
            return rgb_color.is_dark()
        except Exception as err:
            self._log.exception(f"Error getting document color scheme {err}.")
            return True

    @property
    def actual_theme_color_kind(self) -> ThemeColorKind:
        """
        Get the actual theme color kind.

        On MacOS and Windows this will be the same as the ``theme_color_kind``.
        On Linux this may be different when LibreOffice is set to use system color scheme.

        This is because Linux does not always return the correct color scheme when LibreOffice is set to use the system color scheme.
        For this reason the color scheme is determined by the actual document color scheme.
        """
        return self.get_actual_theme_color_kind()

    @property
    def theme_name(self) -> str:
        """Get theme name"""
        return self._theme_name

    @property
    def theme_color_kind(self) -> ThemeColorKind:
        """
        This is the theme color kind that may not be the actual color kind being used.

        Due to inconsistencies in the color scheme detection on Linux for System Theme this may not always be the actual color scheme reported in the configuration.

        If the current LibreOffice Theme is not set to System then this will be the same as the actual color scheme.
        """

        return self.get_theme_color_kind()
