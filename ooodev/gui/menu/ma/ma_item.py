from __future__ import annotations
from typing import List
from ooodev.gui.menu.common.command import Command
from ooodev.gui.menu.common.shortcut import Shortcut
from ooodev.utils.kind.item_style_kind import ItemStyleKind


class MAItem:
    """Class for menu item data."""

    def __init__(
        self,
        *,
        label: str,
        command: str | Command = "",
        style: ItemStyleKind = ItemStyleKind.NONE,
        shortcut: Shortcut | str = "",
        submenu: List[MAItem] | None = None,
    ) -> None:
        self._label = label
        self._command = command
        self._style = style
        self._shortcut = shortcut
        if submenu is None:
            self._submenu = []
        else:
            self._submenu = submenu

    # region Methods
    def is_separator(self) -> bool:
        """Check if menu item is a separator"""
        return self._label == "-"

    def to_dict(self) -> dict:
        """Convert to dictionary"""
        d = {
            "Label": self._label,
            "CommandURL": self._command,
            "Style": self._style,
            "Submenu": [item.to_dict() for item in self._submenu],
        }
        if self._shortcut:
            d["ShortCut"] = self._shortcut
        return d

    # endregion Methods

    # region Properties
    @property
    def label(self) -> str:
        return self._label

    @label.setter
    def label(self, value: str) -> None:
        self._label = value

    @property
    def command(self) -> str | Command:
        return self._command

    @command.setter
    def command(self, value: str | Command) -> None:
        self._command = value

    @property
    def style(self) -> ItemStyleKind:
        return self._style

    @style.setter
    def style(self, value: ItemStyleKind) -> None:
        self._style = value

    @property
    def shortcut(self) -> Shortcut | str:
        return self._shortcut

    @shortcut.setter
    def shortcut(self, value: Shortcut | str) -> None:
        self._shortcut = value

    @property
    def submenu(self) -> List[MAItem]:
        return self._submenu

    @submenu.setter
    def submenu(self, value: List[MAItem]) -> None:
        self._submenu = value

    # endregion Properties
