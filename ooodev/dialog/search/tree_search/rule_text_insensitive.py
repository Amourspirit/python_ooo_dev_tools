from __future__ import annotations
from typing import Any, TYPE_CHECKING


if TYPE_CHECKING:
    from com.sun.star.awt.tree import XTreeNode


class RuleTextInsensitive:
    """Rule for matching a node's text value with a case-insensitive string."""

    def is_match(self, node: XTreeNode, match_value: Any) -> bool:
        """
        Gets the node's text value and compares it with the match value.

        Args:
            node (XTreeNode): Tree node to check.
            match_value (Any): Value to match. Must be a string to match.

        Returns:
            bool: True if the node's text value matches the match value; Otherwise, False.
        """
        node_text = node.getDisplayValue()
        if not node_text:
            return False
        if not isinstance(node_text, str):
            return False
        if not match_value:
            return False
        if not isinstance(match_value, str):
            return False
        return node_text.casefold() == match_value.casefold()
