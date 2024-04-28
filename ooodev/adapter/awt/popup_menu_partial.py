from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.awt import XPopupMenu
from ooodev.adapter.awt.menu_partial import MenuPartial


if TYPE_CHECKING:
    from com.sun.star.awt import XWindowPeer
    from com.sun.star.awt import Rectangle
    from com.sun.star.awt import KeyEvent
    from com.sun.star.graphic import XGraphic
    from ooo.dyn.awt.popup_menu_direction import PopupMenuDirectionEnum
    from ooodev.utils.type_var import UnoInterface
else:
    UnoInterface = Any


class PopupMenuPartial(MenuPartial):
    """
    Partial Class for XPopupMenu.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XPopupMenu, interface: UnoInterface | None = XPopupMenu) -> None:
        """
        Constructor

        Args:
            component (XPopupMenu): UNO Component that implements ``com.sun.star.awt.XPopupMenu`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XPopupMenu``.
        """
        MenuPartial.__init__(self, component=component, interface=interface)
        self.__component = component

    # region XPopupMenu
    def check_item(self, menu_id: int, check: bool = True) -> None:
        """
        Sets the state of the item to be checked or unchecked.

        Args:
            menu_id (int): The item identifier.
            check (bool, optional): The state of the item. Defaults to ``True``.
        """
        self.__component.checkItem(menu_id, check)

    def end_execute(self) -> None:
        """
        Ends the execution of the PopupMenu.

        ``XPopupMenu.execute()`` will then return ``0``.
        """
        self.__component.endExecute()

    def execute(self, parent: XWindowPeer, position: Rectangle, direction: int | PopupMenuDirectionEnum = 0) -> int:
        """
        Executes the popup menu and returns the selected item or 0, if cancelled.

        Args:
            parent (XWindowPeer): The parent window.
            position (Rectangle): The position of the popup menu.
            direction (int | PopupMenuDirectionEnum, optional): The direction of the popup menu. Defaults to ``0``.

        Returns:
            int: The selected item or ``0`` if cancelled.

        Hint:
            - ``PopupMenuDirectionEnum`` is an enum that can be imported from ``ooo.dyn.awt.popup_menu_direction``.

        Note:
            ``direction`` values:

            - ``EXECUTE_DEFAULT = 0``
            - ``EXECUTE_DOWN = 1``
            - ``EXECUTE_UP = 2``
            - ``EXECUTE_LEFT = 4``
            - ``EXECUTE_RIGHT = 8``
        """
        return self.__component.execute(parent, position, int(direction))

    def get_accelerator_key_event(self, menu_id: int) -> KeyEvent:
        """
        Gets the KeyEvent for the menu item.

        The KeyEvent is only used as a container to transport the shortcut information, so that in this case ``com.sun.star.lang.EventObject.Source`` is ``None``.
        """
        return self.__component.getAcceleratorKeyEvent(menu_id)

    def get_default_item(self) -> int:
        """
        Gets the menu default item.
        """
        return self.__component.getDefaultItem()

    def get_item_image(self, menu_id: int) -> XGraphic:
        """
        Gets the image for the menu item.
        """
        return self.__component.getItemImage(menu_id)

    def insert_separator(self, item_pos: int) -> None:
        """
        Inserts a separator at the specified position.
        """
        self.__component.insertSeparator(item_pos)

    def is_in_execute(self) -> bool:
        """
        Queries if the PopupMenu is being.

        Returns ``True`` only if the PopupMenu is being executed as a result of invoking`` XPopupMenu.execute()``;
        that is, for a PopupMenu activated by a MenuBar item, this methods returns ``False``.
        """
        return self.__component.isInExecute()

    def is_item_checked(self, menu_id: int) -> bool:
        """
        Gets whether the item is checked or unchecked.
        """
        return self.__component.isItemChecked(menu_id)

    def set_accelerator_key_event(self, menu_id: int, key_event: KeyEvent) -> None:
        """
        Sets the KeyEvent for the menu item.

        The KeyEvent is only used as a container to transport the shortcut information, this methods only draws the text corresponding to this keyboard shortcut.
        The client code is responsible for listening to keyboard events (typically done via ``XUserInputInterception``), and dispatch the respective command.
        """
        self.__component.setAcceleratorKeyEvent(menu_id, key_event)

    def set_default_item(self, menu_id: int) -> None:
        """
        Sets the menu default item.
        """
        self.__component.setDefaultItem(menu_id)

    def set_item_image(self, menu_id: int, graphic: XGraphic, scale: bool) -> None:
        """
        Sets the image for the menu item.
        """
        self.__component.setItemImage(menu_id, graphic, scale)

    # endregion XPopupMenu
