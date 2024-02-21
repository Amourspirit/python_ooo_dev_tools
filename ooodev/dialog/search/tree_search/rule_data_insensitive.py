from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.awt.tree import XMutableTreeNode

from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.awt.tree import XTreeNode


class RuleDataInsensitive:
    """
    Rule for matching a node's data value with a case-insensitive string.

    For case sensitive matching, use :py:class:`~.RuleDataCompare`.
    """

    def is_match(self, node: XTreeNode, match_value: Any) -> bool:
        """
        Gets the node's data value and compares it with the match value.

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
        if not data_value:
            return False
        if not isinstance(data_value, str):
            return False
        return data_value.casefold() == match_value.casefold()
