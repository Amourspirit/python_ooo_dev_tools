from __future__ import annotations
from typing import Any
from dataclasses import dataclass


@dataclass
class PopupItem:
    """Class for individual menu item"""

    text: str
    """Menu Text."""
    menu_id: int
    """Menu ID."""
    command: str
    """Menu Command."""
    style: int
    """Menu Style."""
    checked: bool
    """Menu check state."""
    enabled: bool
    """Menu enable state."""
    default: bool
    """Specifies if menu is default item."""
    help_command: str
    """Help command."""
    help_text: str
    """Help text."""
    tip_help_text: str
    """Tip help text."""
    shortcut: str
    """One or more shortcuts separated by semicolon."""
    index: int = -1
    """Menu index, Defaults to -1."""
    data: Any = None
    """Data associated with menu item. Defaults to ``None``."""

    def is_separator(self) -> bool:
        """Check if item is separator"""
        return self.text == "-"
