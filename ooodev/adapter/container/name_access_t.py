from __future__ import annotations
from typing import Any, Protocol
from .element_access_t import ElementAccessT


class NameAccessT(ElementAccessT, Protocol):
    def has_by_name(self, name: str) -> bool:
        ...

    def get_by_name(self, name: str) -> Any:
        ...
