from __future__ import annotations
from typing import Dict, Any
import uno
from ooo.dyn.awt.menu_item_style import MenuItemStyleEnum

from ooodev.gui.menu.popup.builder.item import Item
from ooodev.utils.kind.module_names_kind import ModuleNamesKind


class PopupItem(Item):
    """Class that represents a popup item for the popup menu builder."""

    # region Constructor
    def __init__(
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
    ) -> None:
        """
        Constructor

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

        Hint:
            - ``MenuItemStyleEnum`` is an enum and can be imported from ``ooo.dyn.awt.menu_item_style``.
            - ``ModuleNamesKind`` is an enum and can be imported from ``ooodev.utils.kind.module_names_kind``.
        """
        super().__init__()
        self._text = text
        self._command = command
        self._style = int(style)
        self._checked = checked
        self._enabled = enabled
        self._default = default
        self._help_text = help_text
        self._help_command = help_command
        self._tip_help_text = tip_help_text
        self._shortcut = shortcut
        self._module = module

    # endregion Constructor

    # region Public Methods

    def to_dict(self) -> Dict[str, Any]:
        """
        Gets the PopupItem as a dictionary

        Returns:
            Dict[str, Any]: The PopupItem as a dictionary
        """
        d = {
            "text": self._text,
            "command": self._command,
            "style": self._style,
            "checked": self._checked,
            "enabled": self._enabled,
            "default": self._default,
            "help_text": self._help_text,
            "help_command": self._help_command,
            "tip_help_text": self._tip_help_text,
            "shortcut": self._shortcut,
        }
        if self._module:
            d["module"] = self._module
        return d

    # endregion Public Methods

    @staticmethod
    def from_dict(data: Dict[str, Any]) -> PopupItem:
        """
        Gets a PopupItem from a dictionary

        Args:
            data (Dict[str, Any]): The dictionary to get the PopupItem from

        Returns:
            PopupItem: The PopupItem
        """
        return PopupItem(
            text=data.get("text", ""),
            command=data.get("command", ""),
            style=data.get("style", 0),
            checked=data.get("checked", False),
            enabled=data.get("enabled", True),
            default=data.get("default", False),
            help_text=data.get("help_text", ""),
            help_command=data.get("help_command", ""),
            tip_help_text=data.get("tip_help_text", ""),
            shortcut=data.get("shortcut", ""),
            module=data.get("module"),
        )

    # region Properties
    @property
    def text(self) -> str:
        """Gets/Sets Text."""
        return self._text

    @text.setter
    def text(self, value: str) -> None:
        self._text = value

    @property
    def command(self) -> str:
        """Gets/Sets Command."""
        return self._command

    @command.setter
    def command(self, value: str) -> None:
        self._command = value

    @property
    def style(self) -> int:
        """Gets/Sets Style."""
        return self._style

    @style.setter
    def style(self, value: int | MenuItemStyleEnum) -> None:
        self._style = int(value)

    @property
    def checked(self) -> bool:
        """Gets/Sets Check state."""
        return self._checked

    @checked.setter
    def checked(self, value: bool) -> None:
        self._checked = value

    @property
    def enabled(self) -> bool:
        """Gets/Sets Enable State."""
        return self._enabled

    @enabled.setter
    def enabled(self, value: bool) -> None:
        self._enabled = value

    @property
    def default(self) -> bool:
        """Gets/Sets if entry is default."""
        return self._default

    @default.setter
    def default(self, value: bool) -> None:
        self._default = value

    @property
    def help_text(self) -> str:
        """Gets/Sets Help Text."""
        return self._help_text

    @help_text.setter
    def help_text(self, value: str) -> None:
        self._help_text = value

    @property
    def help_command(self) -> str:
        """Gets/Sets Help command."""
        return self._help_command

    @help_command.setter
    def help_command(self, value: str) -> None:
        self._help_command = value

    @property
    def tip_help_text(self) -> str:
        """Gets/Sets Tip Help Text."""
        return self._tip_help_text

    @tip_help_text.setter
    def tip_help_text(self, value: str) -> None:
        self._tip_help_text = value

    @property
    def shortcut(self) -> str:
        """Gets/Sets Shortcut."""
        return self._shortcut

    @shortcut.setter
    def shortcut(self, value: str) -> None:
        self._shortcut = value

    @property
    def module(self) -> Any:
        """Gets/Sets Module."""
        return self._module

    @module.setter
    def module(self, value: Any) -> None:
        self._module = value

    # endregion Properties
