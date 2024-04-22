from __future__ import annotations
from typing import Any, Union, TYPE_CHECKING
from dataclasses import dataclass

if TYPE_CHECKING:
    from ooodev.gui.menu.common.command_dict import CommandDict


@dataclass
class PopupItem:
    """Class for individual menu item"""

    text: str
    menu_id: int
    command: "Union[str, CommandDict]"
    style: int
    checked: bool
    enabled: bool
    default: bool
    help_command: str
    help_text: str
    tip_help_text: str
    shortcut: str
    index: int = -1
    data: Any = None

    def is_separator(self) -> bool:
        """Check if item is separator"""
        return self.text == "-"
