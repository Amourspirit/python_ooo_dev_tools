from __future__ import annotations
from re import sub
from typing import Dict, List, Any
import uno
from ooo.dyn.awt.menu_item_style import MenuItemStyleEnum

from ooodev.gui.menu.popup.builder.item import Item
from ooodev.gui.menu.popup.builder.sep_item import SepItem
from ooodev.utils.kind.module_names_kind import ModuleNamesKind
from ooodev.gui.menu.common.command import Command


class BuilderItem(Item):
    """Class that represents a popup item for the popup menu builder."""

    # region Constructor
    def __init__(
        self,
        *,
        text: str = "",
        command: str | Command = "",
        style: int | MenuItemStyleEnum = 0,
        checked: bool = False,
        enabled: bool = True,
        default: bool = False,
        help_text: str = "",
        help_command: str = "",
        tip_help_text: str = "",
        shortcut: str = "",
        module: ModuleNamesKind | str | None = None,
        submenu: List[Item] | None = None,
    ) -> None:
        """
        Constructor

        Args:
            text (str, optional): Text.
            command (str, Command, optional): Command.
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
            - ``Command`` can be imported from ``ooodev.gui.menu.common.command``.
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
        if submenu is None:
            self._submenu = []
        else:
            self._submenu = submenu

    # endregion Constructor

    def _process_sub_menu(self, parent: dict, submenus: list[Item]) -> None:
        """Insert submenu"""

        for i, menu in enumerate(submenus):
            if i == 0:
                if "submenu" not in parent:
                    parent["submenu"] = []
            if isinstance(menu, SepItem):
                parent["submenu"].append(menu.to_dict())
                continue
            elif isinstance(menu, BuilderItem):
                new_parent = menu._to_dict()
                parent["submenu"].append(new_parent)
                if menu.submenu:
                    self._process_sub_menu(new_parent, menu.submenu)

    # region Public Methods

    def _to_dict(self) -> Dict[str, Any]:
        """
        Gets the PopupItem as a dictionary

        Returns:
            Dict[str, Any]: The PopupItem as a dictionary
        """
        if isinstance(self._command, str):
            cmd = self._command
        else:
            cmd = self._command.to_dict()

        d = {
            "text": self._text,
            "command": cmd,
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

    def to_dict(self) -> Dict[str, Any]:
        dict_data = self._to_dict()
        if self._submenu:
            self._process_sub_menu(dict_data, self._submenu)
        return dict_data

    def submenu_insert_separator(self, idx: int) -> SepItem:
        """
        Adds a new separator to the submenu.

        Returns:
            SepItem: The new separator.
        """
        sep = SepItem()
        self._submenu.insert(idx, sep)
        return sep

    def submenu_add_separator(self) -> SepItem:
        """
        Adds a new separator to the submenu.

        Returns:
            SepItem: The new separator.
        """
        sep = SepItem()
        self._submenu.append(sep)
        return sep

    def submenu_insert_popup(
        self,
        idx,
        *,
        text: str = "",
        command: str | Command = "",
        style: int | MenuItemStyleEnum = 0,
        checked: bool = False,
        enabled: bool = True,
        default: bool = False,
        help_text: str = "",
        help_command: str = "",
        tip_help_text: str = "",
        shortcut: str = "",
        module: ModuleNamesKind | str | None = None,
    ) -> BuilderItem:
        """
        Inserts a new popup item to the submenu.

        Args:
            idx (int): Index.
            text (str, optional): Text.
            command (str, Command, optional): Command.
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
            PopupItem: The new popup item.

        Hint:
            - ``MenuItemStyleEnum`` is an enum and can be imported from ``ooo.dyn.awt.menu_item_style``.
            - ``ModuleNamesKind`` is an enum and can be imported from ``ooodev.utils.kind.module_names_kind``.
            - ``Command`` can be imported from ``ooodev.gui.menu.common.command``.
        """
        popup = self._create_popup(
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
        self._submenu.insert(idx, popup)
        return popup

    def submenu_add_popup(
        self,
        *,
        text: str = "",
        command: str | Command = "",
        style: int | MenuItemStyleEnum = 0,
        checked: bool = False,
        enabled: bool = True,
        default: bool = False,
        help_text: str = "",
        help_command: str = "",
        tip_help_text: str = "",
        shortcut: str = "",
        module: ModuleNamesKind | str | None = None,
    ) -> BuilderItem:
        """
        Adds a new popup item to the submenu.

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
            PopupItem: The new popup item.

        Hint:
            - ``MenuItemStyleEnum`` is an enum and can be imported from ``ooo.dyn.awt.menu_item_style``.
            - ``ModuleNamesKind`` is an enum and can be imported from ``ooodev.utils.kind.module_names_kind``.
            - ``Command`` can be imported from ``ooodev.gui.menu.common.command``.
        """
        popup = self._create_popup(
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
        self._submenu.append(popup)
        return popup

    def _create_popup(
        self,
        *,
        text: str = "",
        command: str | Command = "",
        style: int | MenuItemStyleEnum = 0,
        checked: bool = False,
        enabled: bool = True,
        default: bool = False,
        help_text: str = "",
        help_command: str = "",
        tip_help_text: str = "",
        shortcut: str = "",
        module: ModuleNamesKind | str | None = None,
    ) -> BuilderItem:
        popup = BuilderItem(
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
        return popup

    # endregion Public Methods

    # region Properties
    @property
    def text(self) -> str:
        """Gets/Sets Text."""
        return self._text

    @text.setter
    def text(self, value: str) -> None:
        self._text = value

    @property
    def command(self) -> str | Command:
        """Gets/Sets Command."""
        return self._command

    @command.setter
    def command(self, value: str | Command) -> None:
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

    @property
    def submenu(self) -> list[Item]:
        """Gets Submenu."""
        return self._submenu

    # endregion Properties
