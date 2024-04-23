from __future__ import annotations
from typing import Dict, Any
from abc import ABC, abstractmethod


class Item(ABC):

    @abstractmethod
    def to_dict(self) -> Dict[str, Any]:
        raise NotImplementedError

    @property
    def is_separator(self) -> bool:
        """Gets if item is a separator."""
        return False
