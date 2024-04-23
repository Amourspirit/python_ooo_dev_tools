from __future__ import annotations
from typing import Dict
from ooodev.gui.menu.popup.builder.item import Item


class SepItem(Item):
    def __init__(self) -> None:
        super().__init__()

    def to_dict(self) -> Dict[str, str]:
        return {"text": "-"}

    @property
    def is_separator(self) -> bool:
        """Gets if item is a separator."""
        return True
