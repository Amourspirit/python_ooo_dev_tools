from __future__ import annotations
from dataclasses import dataclass, asdict


@dataclass
class Command:
    name: str
    language: str = "Basic"
    location: str = "user"
    library: str = "standard"
    module: str = "."

    def to_dict(self) -> dict:
        return asdict(self)
