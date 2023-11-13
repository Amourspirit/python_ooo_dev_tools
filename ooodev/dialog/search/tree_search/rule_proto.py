from __future__ import annotations
from typing import Any, Protocol, TYPE_CHECKING

if TYPE_CHECKING:
    from com.sun.star.awt.tree import XTreeNode


class RuleT(Protocol):
    def is_match(self, node: XTreeNode, match_value: Any) -> bool:
        ...
