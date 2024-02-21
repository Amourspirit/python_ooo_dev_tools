from __future__ import annotations
from typing import Any, Literal, TYPE_CHECKING
import uno
from com.sun.star.awt.tree import XMutableTreeNode

from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.awt.tree import XTreeNode


class RuleDataCompare:
    """
    Rule for matching a node's data value match value.

    This rule can also be used for case-sensitive matching by setting ``compare_op = "="``.
    """

    def __init__(self, compare_op: Literal["<", ">", "=", "<=", ">=", "!="]) -> None:
        """
        Constructor

        Args:
            compare_op (Literal[, optional): which compare operation to preform. Defaults to "].
        """
        if compare_op == "<":
            self._compare_op = -1
        elif compare_op == ">":
            self._compare_op = 1
        elif compare_op == "!=":
            self._compare_op = 3
        elif compare_op == "<=":
            self._compare_op = -2
        elif compare_op == ">=":
            self._compare_op = 2
        else:
            # equals
            self._compare_op = 0

    def _compare(self, data_value: Any, match_value: Any) -> bool:
        if self._compare_op == -1:
            return data_value < match_value
        elif self._compare_op == 1:
            return data_value > match_value
        elif self._compare_op == 3:
            return data_value != match_value
        elif self._compare_op == -2:
            return data_value <= match_value
        elif self._compare_op == 2:
            return data_value >= match_value
        else:
            # equals
            return data_value == match_value

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
        if data_value is None:
            if match_value is None:
                return True
            return False
        return self._compare(data_value, match_value)
