from __future__ import annotations
from typing import Any, TYPE_CHECKING


if TYPE_CHECKING:
    try:
        from typing import Protocol
    except ImportError:
        from typing_extensions import Protocol
else:
    Protocol = object


class CallableAnyT(Protocol):
    def __call__(self, *args: Any, **kwargs: Any) -> Any: ...
