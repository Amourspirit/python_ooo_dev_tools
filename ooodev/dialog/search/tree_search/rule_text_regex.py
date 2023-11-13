from __future__ import annotations
from typing import Any, TYPE_CHECKING
import re

if TYPE_CHECKING:
    from com.sun.star.awt.tree import XTreeNode


class RuleTextRegex:
    """Rule for matching a node's text value with a regular expression pattern."""

    def is_match(self, node: XTreeNode, match_value: Any) -> bool:
        """
        Gets the node's text value and compares it with the match value.

        Args:
            node (XTreeNode): Tree node to check.
            match_value (Any): Value to match. Must be a regular expression pattern to match.

        Returns:
            bool: True if the node's data value matches the match value; Otherwise, False.
        """
        node_text = node.getDisplayValue()
        if not match_value:
            return False
        if not isinstance(match_value, re.Pattern):
            return False
        return match_value.search(node_text) is not None
