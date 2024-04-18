from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.awt import XMenu
from ooo.dyn.awt.menu_item_type import MenuItemType

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.awt import XMenuListener
    from com.sun.star.awt import XPopupMenu
    from ooodev.utils.kind.menu_item_style_kind import MenuItemStyleKind
    from ooodev.utils.type_var import UnoInterface


class MenuPartial:
    """
    Partial class for XMenu.
    """

    def __init__(self, component: XMenu, interface: UnoInterface | None = XMenu) -> None:
        """
        Constructor

        Args:
            component (XMenu): UNO Component that implements ``com.sun.star.awt.XMenu`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XMenu``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    def __len__(self) -> int:
        """
        Gets the number of items in the menu.
        """
        return self.get_item_count()

    # region XMenu
    def add_menu_listener(self, listener: XMenuListener) -> None:
        """
        adds the specified menu listener to receive events from this menu.
        """
        self.__component.addMenuListener(listener)

    def clear(self) -> None:
        """
        Removes all items from the menu.
        """
        self.__component.clear()

    def enable_auto_mnemonics(self, enable: bool) -> None:
        """
        Specifies whether mnemonics are automatically assigned to menu items, or not.
        """
        self.__component.enableAutoMnemonics(enable)

    def enable_item(self, menu_id: int, enable: bool = False) -> None:
        """
        Enables or disables the menu item.

        Args:
            menu_id (int): The ID of the menu item.
            enable (bool, optional): The state of the menu item. Defaults to ``False``.
        """
        self.__component.enableItem(menu_id, enable)

    def get_command(self, menu_id: int) -> str:
        """
        Gets the command string for the menu item.
        """
        return self.__component.getCommand(menu_id)

    def get_help_command(self, menu_id: int) -> str:
        """
        Gets the help command string for the menu item.
        """
        return self.__component.getHelpCommand(menu_id)

    def get_help_text(self, menu_id: int) -> str:
        """
        Gets the help text for the menu item.
        """
        return self.__component.getHelpText(menu_id)

    def get_item_count(self) -> int:
        """
        Gets the number of items in the menu.
        """
        return self.__component.getItemCount()

    def get_item_id(self, item_pos: int) -> int:
        """
        Gets the ID of the item at the specified position.
        """
        return self.__component.getItemId(item_pos)

    def get_item_pos(self, menu_id: int) -> int:
        """
        Gets the position of the item with the specified ID.
        """
        return self.__component.getItemPos(menu_id)

    def get_item_text(self, menu_id: int) -> str:
        """
        returns the string for the given item id.
        """
        return self.__component.getItemText(menu_id)

    def get_item_type(self, item_pos: int) -> MenuItemType:
        """
        Gets the type of the menu item.

        Args:
            item_pos (int): The position of the menu item.

        Returns:
            MenuItemType: The type of the menu item.

        Hint:
            - ``MenuItemType`` is an enum and can be imported from ``ooo.dyn.awt.menu_item_type``.
        """
        return self.__component.getItemType(item_pos)  # type: ignore

    def get_popup_menu(self, menu_id: int) -> XPopupMenu:
        """
        Gets the popup menu from the menu item.
        """
        return self.__component.getPopupMenu(menu_id)

    def get_tip_help_text(self, item_id: int) -> str:
        """
        retrieves the tip help text for the menu item.
        """
        return self.__component.getTipHelpText(item_id)

    def hide_disabled_entries(self, hide: bool = True) -> None:
        """
        Specifies whether disabled menu entries should be hidden, or not.

        Args:
            hide (bool, optional): ``True`` to hide disabled entries, ``False`` to show them. Defaults to ``True``.
        """
        self.__component.hideDisabledEntries(hide)

    def insert_item(self, menu_id: int, text: str, item_style: int | MenuItemStyleKind, item_pos: int) -> None:
        """
        Inserts an item into the menu.

        The item is appended if the position is greater than or equal to ``get_item_count()`` or if it is negative.

        Args:
            menu_id (int): The ID of the menu item.
            text (str): The text for the menu item.
            item_style (int | MenuItemStyleKind): The style of the menu item. ``MenuItemStyleKind`` is a flag enum.
            item_pos (int): The position of the menu item.

        Returns:
            None:

        Hint:
            - ``MenuItemStyleKind`` is an enum and can be imported from ``ooodev.utils.kind.menu_item_style_kind``.
        """
        self.__component.insertItem(menu_id, text, int(item_style), item_pos)

    def is_item_enabled(self, menu_id: int) -> bool:
        """
        Gets the state of the menu item.
        """
        return self.__component.isItemEnabled(menu_id)

    def is_popup_menu(self) -> bool:
        """
        Checks whether an ``XMenu`` is an ``XPopupMenu``.
        """
        return self.__component.isPopupMenu()

    def remove_item(self, item_pos: int, count: int) -> None:
        """
        Removes one or more items from the menu.
        """
        self.__component.removeItem(item_pos, count)

    def remove_menu_listener(self, listener: XMenuListener) -> None:
        """
        Removes the specified menu listener so that it no longer receives events from this menu.
        """
        self.__component.removeMenuListener(listener)

    def set_command(self, menu_id: int, command: str) -> None:
        """
        Sets the command string for the menu item.
        """
        self.__component.setCommand(menu_id, command)

    def set_help_command(self, menu_id: int, command: str) -> None:
        """
        Sets the help command string for the menu item.
        """
        self.__component.setHelpCommand(menu_id, command)

    def set_help_text(self, menu_id: int, text: str) -> None:
        """
        Sets the help text for the menu item.
        """
        self.__component.setHelpText(menu_id, text)

    def set_item_text(self, menu_id: int, text: str) -> None:
        """
        Sets the text for the menu item.
        """
        self.__component.setItemText(menu_id, text)

    def set_popup_menu(self, menu_id: int, popup_menu: XPopupMenu) -> None:
        """
        Sets the popup menu for a specified menu item.
        """
        self.__component.setPopupMenu(menu_id, popup_menu)

    def set_tip_help_text(self, menu_id: int, text: str) -> None:
        """
        Sets the tip help text for the menu item.
        """
        self.__component.setTipHelpText(menu_id, text)

    # endregion XMenu
