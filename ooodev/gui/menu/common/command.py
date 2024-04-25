from __future__ import annotations
from typing import TYPE_CHECKING
from dataclasses import dataclass, asdict

if TYPE_CHECKING:
    from ooodev.gui.menu.common.command_dict import CommandDict


@dataclass
class Command:
    name: str
    language: str = "Basic"
    location: str = "user"
    library: str = "standard"
    module: str = "."

    def to_dict(self) -> CommandDict:
        return asdict(self)  # type: ignore

    def __bool__(self) -> bool:
        return bool(self.name)
