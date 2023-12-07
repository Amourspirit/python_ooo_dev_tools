from __future__ import annotations
from typing import Any, TYPE_CHECKING, Union

if TYPE_CHECKING:
    try:
        from typing import Protocol
    except ImportError:
        from typing_extensions import Protocol
else:
    Protocol = object


class ComponentStructT(Protocol):
    component: Any


class ComponentClassT(Protocol):
    @property
    def component(self) -> Any:
        raise NotImplementedError()


ComponentT = Union[ComponentStructT, ComponentClassT]
