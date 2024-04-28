from __future__ import annotations
from typing import List

from ooodev.gui.menu.ma.ma_creator import MACreator
from ooodev.gui.menu.popup.popup_creator import PopupCreator
from ooodev.gui.menu.context.context_creator import ContextCreator


class MenuConvert:
    """Class for converting menu data."""

    @staticmethod
    def convert_from_app_menu_to_popup(menus: List[dict[str, str]]) -> List[dict[str, str]]:
        """
        Convert app menu to popup menu.

        Args:
            menus (List[dict[str, str]]): App menu data.

        Returns:
            List[dict[str, str]]: Popup menu data.
        """
        lookups = {
            "CommandURL": "command",
            "Label": "text",
            "ShortCut": "shortcut",
            "Style": "style",
            "Submenu": "submenu",
        }

        creator = MACreator()
        creator.key_lookups = lookups
        json_dict = creator.get_json_dict(menus)
        return json_dict

    @staticmethod
    def convert_from_app_menu_to_context(menus: List[dict[str, str]]) -> List[dict[str, str]]:
        """
        Convert app menu to popup menu.

        Args:
            menus (List[dict[str, str]]): App menu data.

        Returns:
            List[dict[str, str]]: Popup menu data.
        """
        lookups = {
            "CommandURL": "command",
            "Label": "text",
            "Style": "style",
            "Submenu": "submenu",
        }

        creator = MACreator()
        creator.key_lookups = lookups
        json_dict = creator.get_json_dict(menus)
        return json_dict

    @staticmethod
    def convert_from_popup_to_app_menu(menus: List[dict[str, str]]) -> List[dict[str, str]]:
        """
        Convert popup menu to app menu.

        Args:
            menus (List[dict[str, str]]): Popup menu data.

        Returns:
            List[dict[str, str]]: App menu data.
        """
        lookups = {
            "command": "CommandURL",
            "text": "Label",
            "shortcut": "ShortCut",
            "style": "Style",
            "submenu": "Submenu",
        }

        creator = PopupCreator()
        creator.key_lookups = lookups
        json_dict = creator.get_json_dict(menus)
        return json_dict

    @staticmethod
    def convert_from_popup_context_menu(menus: List[dict[str, str]]) -> List[dict[str, str]]:
        """
        Convert popup menu to context action menu.

        Args:
            menus (List[dict[str, str]]): Popup menu data.

        Returns:
            List[dict[str, str]]: Context Action menu data.
        """
        lookups = {
            "command": "command",
            "text": "text",
            "submenu": "Submenu",
        }

        creator = PopupCreator()
        creator.key_lookups = lookups
        json_dict = creator.get_json_dict(menus)
        return json_dict

    @staticmethod
    def convert_from_context_to_app_menu(menus: List[dict[str, str]]) -> List[dict[str, str]]:
        """
        Convert context action menu to app menu.

        Args:
            menus (List[dict[str, str]]): Popup menu data.

        Returns:
            List[dict[str, str]]: App menu data.
        """
        lookups = {
            "command": "CommandURL",
            "text": "Label",
            "submenu": "Submenu",
        }

        creator = ContextCreator()
        creator.key_lookups = lookups
        json_dict = creator.get_json_dict(menus)
        return json_dict

    @staticmethod
    def convert_from_context_to_popup_menu(menus: List[dict[str, str]]) -> List[dict[str, str]]:
        """
        Convert context action menu to popup menu.

        Args:
            menus (List[dict[str, str]]): Context Action menu data.

        Returns:
            List[dict[str, str]]: Popup menu data.
        """
        lookups = {
            "command": "command",
            "text": "text",
            "help_command": "help_command",
            "submenu": "submenu",
        }

        creator = ContextCreator()
        creator.key_lookups = lookups
        json_dict = creator.get_json_dict(menus)
        return json_dict
