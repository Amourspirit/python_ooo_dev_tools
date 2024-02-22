from __future__ import annotations
from typing import Any, TYPE_CHECKING
import contextlib
import uno
from com.sun.star.awt.tree import XMutableTreeNode

from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.awt.tree import XTreeNode


class RuleDataInstance:
    """Rule for matching a node's data type with a specified type."""

    def is_match(self, node: XTreeNode, match_value: Any) -> bool:
        """
        Gets the node's data value and compares it with the match value for instance match.

        Args:
            node (XTreeNode): Tree node to check.
            match_value (Any): Value to match. Must be a string to match.

        Returns:
            bool: True if the node's data value matches the match value; Otherwise, False.
        """
        tree_node = mLo.Lo.qi(XMutableTreeNode, node)
        if not tree_node:
            return False
        data_value = tree_node.DataValue
        if data_value is None:
            return False
        with contextlib.suppress(Exception):
            return isinstance(data_value, match_value)
        return False
