from __future__ import annotations
from typing import Any, Protocol
from .index_access_t import IndexAccessT


class IndexReplaceT(IndexAccessT, Protocol):
    def replace_by_index(self, index: int, element: Any) -> None:
        ...
