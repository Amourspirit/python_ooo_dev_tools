from __future__ import annotations
from typing import Any, Protocol
from .index_replace_t import IndexReplaceT


class IndexContainerT(IndexReplaceT, Protocol):
    def insert_by_index(self, index: int, element: Any) -> None:
        ...

    def remove_by_index(self, index: int) -> None:
        ...
