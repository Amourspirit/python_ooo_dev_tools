from __future__ import annotations
from typing import Dict, Any, TYPE_CHECKING
import uno
from ooo.dyn.awt.menu_item_style import MenuItemStyleEnum

from ooodev.utils.kind.module_names_kind import ModuleNamesKind
from ooodev.gui.menu.popup.builder.popup_item import PopupItem
from ooodev.gui.menu.popup.builder.sep_item import SepItem

if TYPE_CHECKING:
    from ooodev.gui.menu.popup.builder.item import Item


class PopupBuilder:
    def __init__(self) -> None:
        self._menus = []

    # region Dunder Methods
    def __len__(self) -> int:
        """Gets the length of the menus."""
        return len(self._menus)

    def __getitem__(self, index: int) -> Item:
        """
        Gets an item at the specified index.

        Args:
            index (int): Index.

        Returns:
            Item: Item.
        """
        return self._menus[index]

    def __setitem__(self, index: int, value: Item) -> None:
        """
        Sets an item at the specified index.

        Args:
            index (int): Index.
            value (Item): Item.
        """
        self._menus[index] = value

    def __delitem__(self, index: int) -> None:
        """
        Deletes an item at the specified index.

        Args:
            index (int): Index.
        """
        del self._menus[index]

    def __iter__(self) -> Any:
        """Iterates over the menu"""
        return iter(self._menus)

    # endregion Dunder Methods

    # region Public Methods

    def insert(self, index: int, item: Item) -> PopupBuilder:
        """Inserts an item at the specified index."""
        self._menus.insert(index, item)
        return self

    def remove(self, item: Item) -> PopupBuilder:
        """Removes an item."""
        self._menus.remove(item)
        return self

    def clear(self) -> PopupBuilder:
        """Clears all items."""
        self._menus.clear()
        return self

    def remove_at_index(self, index: int) -> PopupBuilder:
        """Removes an item at the specified index."""
        del self._menus[index]
        return self

    def pop(self, index: int) -> Item:
        """Pops an item at the specified index."""
        return self._menus.pop(index)

    def add_separator(self) -> PopupBuilder:
        self._menus.append(SepItem())
        return self

    def add_item(self, item: Item) -> PopupBuilder:
        self._menus.append(item)
        return self

    def add(
        self,
        *,
        text: str = "",
        command: str = "",
        style: int | MenuItemStyleEnum = 0,
        checked: bool = False,
        enabled: bool = True,
        default: bool = False,
        help_text: str = "",
        help_command: str = "",
        tip_help_text: str = "",
        shortcut: str = "",
        module: ModuleNamesKind | str | None = None,
    ) -> PopupBuilder:
        """
        Add a popup item.

        Args:
            text (str, optional): Text.
            command (str, optional): Command.
            style (int, MenuItemStyleEnum, optional): Style. Defaults to ``0``.
            checked (bool, optional): Check state. Defaults to ``False``.
            enabled (bool, optional): Enable State. Defaults to ``True``.
            default (bool, optional): Specifies if entry is default. Defaults to ``False``.
            help_text (str, optional): Help Text.
            help_command (str, optional): Help command.
            tip_help_text (str, optional): Tip Help Text.
            shortcut (str, optional): Shortcut.
            module (Any, optional): Module. Defaults to ``None``.

        Returns:
            PopupBuilder: PopupBuilder instance.
        """
        if text == "-":
            return self.add_separator()
        self._menus.append(
            PopupItem(
                text=text,
                command=command,
                style=style,
                checked=checked,
                enabled=enabled,
                default=default,
                help_text=help_text,
                help_command=help_command,
                tip_help_text=tip_help_text,
                shortcut=shortcut,
                module=module,
            )
        )
        return self

    def build(self) -> list[Dict[str, Any]]:
        """Gets the menus as a list of dictionaries."""
        result = []
        for menu in self._menus:
            result.append(menu.to_dict())
        return result

    def get_last(self) -> Item:
        """Gets the last item."""
        if not self._menus:
            raise ValueError("No items found.")
        return self._menus[-1]

    def move_up(self, index: int) -> PopupBuilder:
        """Moves an item up by one position."""
        if index > 0:
            self._menus[index], self._menus[index - 1] = self._menus[index - 1], self._menus[index]
        return self

    def move_down(self, index: int) -> PopupBuilder:
        """Moves an item down by one position."""
        if index < len(self._menus) - 1:
            self._menus[index], self._menus[index + 1] = self._menus[index + 1], self._menus[index]
        return self

    # endregion Public Methods
