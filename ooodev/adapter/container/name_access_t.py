from __future__ import annotations
from typing import Any, TYPE_CHECKING

if TYPE_CHECKING:
    from typing_extensions import Protocol
    from ooodev.adapter.container.element_access_t import ElementAccessT

    class NameAccessT(ElementAccessT, Protocol):
        def has_by_name(self, name: str) -> bool: ...

        def get_by_name(self, name: str) -> Any: ...

else:
    NameAccessT = object
