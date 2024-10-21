from __future__ import annotations
from typing import Dict

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

from ooodev.gui.menu.popup.builder.item import Item


class SepItem(Item):
    def __init__(self) -> None:
        super().__init__()

    @override
    def to_dict(self) -> Dict[str, str]:
        return {"text": "-"}

    @property
    @override
    def is_separator(self) -> bool:
        """Gets if item is a separator."""
        return True
