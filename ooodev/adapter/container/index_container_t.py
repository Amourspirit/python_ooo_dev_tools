from __future__ import annotations
from typing import Any, TYPE_CHECKING

if TYPE_CHECKING:
    from typing_extensions import Protocol
    from ooodev.adapter.container.index_replace_t import IndexReplaceT

    class IndexContainerT(IndexReplaceT, Protocol):
        def insert_by_index(self, index: int, element: Any) -> None: ...

        def remove_by_index(self, index: int) -> None: ...

else:
    IndexContainerT = object
