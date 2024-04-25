from __future__ import annotations
from dataclasses import dataclass, asdict


@dataclass
class Shortcut:
    key: str
    save: bool = True

    def to_dict(self) -> dict:
        return asdict(self)
