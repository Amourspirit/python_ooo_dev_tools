from __future__ import annotations
from typing import Any, TYPE_CHECKING

if TYPE_CHECKING:
    from typing_extensions import Protocol
    from ooodev.adapter.container.index_access_t import IndexAccessT

    class IndexReplaceT(IndexAccessT, Protocol):
        def replace_by_index(self, index: int, element: Any) -> None: ...

else:
    IndexReplaceT = object
