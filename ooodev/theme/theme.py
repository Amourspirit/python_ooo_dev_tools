from ..utils.info import Info
from abc import ABC
from enum import Enum


class ThemeKind(Enum):
    """Theme Kind Lookup"""

    LIBRE_OFFICE = "LibreOffice"
    LIBRE_OFFICE_DARK = "LibreOffice Dark"

    def __str__(self) -> str:
        return self.value


class ThemeBase(ABC):
    def __init__(self, theme_name: ThemeKind | str = "") -> None:
        """
        Constructor

        Args:
            theme_name (ThemeKind | str, optional): Theme Name. If omitted then the current LibreOffice Theme is used.

        Returns:
            None:
        """
        if theme_name == "":
            theme_name = Info.get_office_theme()
        if not theme_name:
            raise ValueError("No theme name has been found,")
        self._theme_name = str(theme_name)

    def _get_color(self, prop_name: str) -> int:
        val = Info.get_config(
            node_str="Color",
            node_path=f"org.openoffice.Office.UI/ColorScheme/ColorSchemes/org.openoffice.Office.UI:ColorScheme['{self._theme_name}']/{prop_name}",
        )
        if val is None:
            return -1
        return int(val)

    def _get_visible(self, prop_name: str) -> bool:
        val = Info.get_config(
            node_str="IsVisible",
            node_path=f"org.openoffice.Office.UI/ColorScheme/ColorSchemes/org.openoffice.Office.UI:ColorScheme['{self._theme_name}']/{prop_name}",
        )
        if val is None:
            return False
        return bool(val)

    @property
    def theme_name(self) -> str:
        """Get theme name"""
        return self._theme_name
