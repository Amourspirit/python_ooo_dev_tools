from __future__ import annotations
from typing import Any, TYPE_CHECKING

if TYPE_CHECKING:
    from typing_extensions import Protocol
else:
    Protocol = object


class ElementAccessT(Protocol):
    def get_element_type(self) -> Any: ...

    def has_elements(self) -> bool: ...
