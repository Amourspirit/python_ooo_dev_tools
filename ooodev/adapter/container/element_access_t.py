from __future__ import annotations
from typing import Any, Protocol


class ElementAccessT(Protocol):
    def get_element_type(self) -> Any:
        ...

    def has_elements(self) -> bool:
        ...
