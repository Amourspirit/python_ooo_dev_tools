from __future__ import annotations
from typing import Any, Protocol
from .element_access_t import ElementAccessT


class IndexAccessT(ElementAccessT, Protocol):
    def get_count(self) -> int:
        ...

    def get_by_index(self, idx: int) -> Any:
        ...
