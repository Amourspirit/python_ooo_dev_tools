from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno

from com.sun.star.awt import XListBox

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.awt import XActionListener, XItemListener
    from ooodev.utils.type_var import UnoInterface


class ListBoxPartial:
    """
    Partial class for XListBox.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XListBox, interface: UnoInterface | None = XListBox) -> None:
        """
        Constructor

        Args:
            component (XListBox): UNO Component that implements ``com.sun.star.awt.XListBox`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XListBox``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XListBox
    def add_action_listener(self, listener: XActionListener) -> None:
        """
        Registers a listener for action events.
        """
        self.__component.addActionListener(listener)

    def add_item(self, item: str, pos: int) -> None:
        """
        Adds an item at the specified position.
        """
        self.__component.addItem(item, pos)

    def add_item_listener(self, listener: XItemListener) -> None:
        """
        Registers a listener for item events.
        """
        self.__component.addItemListener(listener)

    def add_items(self, pos: int, *items: str) -> None:
        """
        Adds multiple items at the specified position.
        """
        self.__component.addItems(items, pos)

    def get_drop_down_line_count(self) -> int:
        """
        Gets the number of visible lines in drop down mode.
        """
        return self.__component.getDropDownLineCount()

    def get_item(self, pos: int) -> str:
        """
        Gets the item at the specified position.
        """
        return self.__component.getItem(pos)

    def get_item_count(self) -> int:
        """
        Gets the number of items in the listbox.
        """
        return self.__component.getItemCount()

    def get_items(self) -> Tuple[str, ...]:
        """
        Gets all items of the list box.
        """
        return self.__component.getItems()

    def get_selected_item(self) -> str:
        """
        Gets the currently selected item.

        When multiple items are selected, the first one is returned. When nothing is selected, an empty string is returned.
        """
        return self.__component.getSelectedItem()

    def get_selected_item_pos(self) -> int:
        """
        GEts the position of the currently selected item.

        When multiple items are selected, the position of the first one is returned. When nothing is selected, -1 is returned.
        """
        return self.__component.getSelectedItemPos()

    def get_selected_items(self) -> Tuple[str, ...]:
        """
        Gets all currently selected items.
        """
        return self.__component.getSelectedItems()

    def get_selected_items_pos(self) -> Tuple[int, ...]:
        """
        Gets the positions of all currently selected items.
        """
        return self.__component.getSelectedItemsPos()  # type: ignore

    def is_multiple_mode(self) -> bool:
        """
        Gets - returns ``True`` if multiple items can be selected, ``False`` if only one item can be selected.
        """
        return self.__component.isMutipleMode()

    def make_visible(self, entry: int) -> None:
        """
        Makes the item at the specified position visible by scrolling.
        """
        self.__component.makeVisible(entry)

    def remove_action_listener(self, listener: XActionListener) -> None:
        """
        Un-registers a listener for action events.
        """
        self.__component.removeActionListener(listener)

    def remove_item_listener(self, listener: XItemListener) -> None:
        """
        Un-registers a listener for item events.
        """
        self.__component.removeItemListener(listener)

    def remove_items(self, pos: int, count: int) -> None:
        """
        Removes a number of items at the specified position.
        """
        self.__component.removeItems(pos, count)

    def select_item(self, item: str, select: bool) -> None:
        """
        Selects/deselects the specified item.
        """
        self.__component.selectItem(item, select)

    def select_item_pos(self, pos: int, select: bool) -> None:
        """
        Selects/deselects the item at the specified position.
        """
        self.__component.selectItemPos(pos, select)

    def select_items_pos(self, select: bool, *positions: int) -> None:
        """
        selects/deselects multiple items at the specified positions.
        """
        self.__component.selectItemsPos(positions, select)  # type: ignore

    def set_drop_down_line_count(self, lines: int) -> None:
        """
        Sets the number of visible lines for drop down mode.
        """
        self.__component.setDropDownLineCount(lines)

    def set_multiple_mode(self, multi: bool) -> None:
        """
        Determines if only a single item or multiple items can be selected.
        """
        self.__component.setMultipleMode(multi)

    # endregion XListBox
