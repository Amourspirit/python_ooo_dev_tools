from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno

from com.sun.star.awt import XComboBox
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.awt import XActionListener
    from com.sun.star.awt import XItemListener
    from ooodev.utils.type_var import UnoInterface


class ComboBoxPartial:
    """
    Partial class for XComboBox.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XComboBox, interface: UnoInterface | None = XComboBox) -> None:
        """
        Constructor

        Args:
            component (XComboBox): UNO Component that implements ``com.sun.star.awt.XComboBox`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XComboBox``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XComboBox
    def add_action_listener(self, listener: XActionListener) -> None:
        """
        registers a listener for action events.
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
        Returns the number of visible lines in the drop down mode.
        """
        return self.__component.getDropDownLineCount()

    def get_item(self, pos: int) -> str:
        """
        Gets the item at the specified position.
        """
        return self.__component.getItem(pos)

    def get_item_count(self) -> int:
        """
        Gets the number of items in the combo box.
        """
        return self.__component.getItemCount()

    def get_items(self) -> Tuple[str, ...]:
        """
        Gets all items of the combo box.
        """
        return self.__component.getItems()

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

    def set_drop_down_line_count(self, lines: int) -> None:
        """
        Sets the number of visible lines for drop down mode.
        """
        self.__component.setDropDownLineCount(lines)

    # endregion XComboBox
