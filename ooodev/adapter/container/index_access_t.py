from __future__ import annotations
from typing import Any, TYPE_CHECKING

if TYPE_CHECKING:
    from typing_extensions import Protocol
    from ooodev.adapter.container.element_access_t import ElementAccessT

    class IndexAccessT(ElementAccessT, Protocol):
        def get_count(self) -> int: ...

        def get_by_index(self, idx: int) -> Any: ...

else:
    IndexAccessT = object
