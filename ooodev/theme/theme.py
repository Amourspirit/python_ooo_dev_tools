from __future__ import annotations
from abc import ABC
from enum import Enum
from ooodev.utils.info import Info


class ThemeKind(Enum):
    """Theme Kind Lookup"""

    AUTOMATIC = "COLOR_SCHEME_LIBREOFFICE_AUTOMATIC"
    LIBRE_OFFICE = "LibreOffice"
    LIBRE_OFFICE_DARK = "LibreOffice Dark"

    def __str__(self) -> str:
        return self.value


class ThemeColorKind(Enum):
    UNKNOWN = -1
    SYSTEM = 0
    LIGHT = 1
    DARK = 2


class ThemeBase(ABC):
    def __init__(self, theme_name: ThemeKind | str = "") -> None:
        """
        Constructor

        Args:
            theme_name (ThemeKind | str, optional): Theme Name. If omitted then the current LibreOffice Theme is used.

        Returns:
            None:
        """
        if theme_name == "" or theme_name == ThemeKind.AUTOMATIC:
            theme_name = Info.get_office_theme()
        if not theme_name:
            raise ValueError("No theme name has been found,")
        self._theme_name = str(theme_name)

    def _get_color(self, prop_name: str) -> int:
        try:
            # val = Info.get_config(
            #     node_str="Color",
            #     node_path=f"org.openoffice.Office.UI/ColorScheme/ColorSchemes/org.openoffice.Office.UI:ColorScheme['{self._theme_name}']/{prop_name}",
            # )
            val = Info.get_config(
                node_str="Color",
                node_path=f"/org.openoffice.Office.UI/ColorScheme/ColorSchemes/{self._theme_name}/{prop_name}",
            )
            return -1 if val is None else int(val)
        except Exception:
            return -1

    def _get_visible(self, prop_name: str) -> bool:
        try:
            # val = Info.get_config(
            #     node_str="IsVisible",
            #     node_path=f"org.openoffice.Office.UI/ColorScheme/ColorSchemes/org.openoffice.Office.UI:ColorScheme['{self._theme_name}']/{prop_name}",
            # )
            val = Info.get_config(
                node_str="IsVisible",
                node_path=f"/org.openoffice.Office.UI/ColorScheme/ColorSchemes/{self._theme_name}/{prop_name}",
            )
            return False if val is None else bool(val)
        except Exception:
            return False

    def get_theme_color_kind(self) -> ThemeColorKind:
        """
        Get Theme Color Kind from configuration.

        Returns:
            ThemeColorKind: Theme Color Kind
        """
        try:
            val = Info.get_config(
                node_str="ApplicationAppearance",
                node_path="org.openoffice.Office.Common/Misc",
            )
            if val == 1:
                return ThemeColorKind.LIGHT
            elif val == 2:
                return ThemeColorKind.DARK
            elif val == 0:
                return ThemeColorKind.SYSTEM
            else:
                return ThemeColorKind.UNKNOWN
        except Exception:
            return ThemeColorKind.UNKNOWN

    @property
    def theme_name(self) -> str:
        """Get theme name"""
        return self._theme_name
