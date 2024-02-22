from __future__ import annotations
from typing import Any, List, TYPE_CHECKING
from .rule_proto import RuleT

if TYPE_CHECKING:
    from com.sun.star.awt.tree import XTreeNode


class SearchTree:
    """Rule engine for searching a tree."""

    def __init__(self, match_value: Any, match_all: bool = False) -> None:
        """
        Constructor
        """
        self._rules: List[RuleT] = []
        self._match_all = match_all
        # self._node = node
        self._match_value = match_value

    def __len__(self) -> int:
        return len(self._rules)

    # region Methods

    def register_rule(self, rule: RuleT) -> None:
        """
        Register rule

        Args:
            rule (RuleT): Rule to register
        """
        if rule in self._rules:
            return
        self._reg_rule(rule=rule)
        self._rules_init = False

    def unregister_rule(self, rule: RuleT):
        """
        Unregister Rule

        Args:
            rule (RuleT): Rule to unregister

        Raises:
            ValueError: If an error occurs
        """
        try:
            self._rules.remove(rule)
            self._rules_init = False
        except ValueError as e:
            msg = f"{self.__class__.__name__}.unregister_rule() Unable to unregister rule."
            raise ValueError(msg) from e

    def _reg_rule(self, rule: RuleT):
        self._rules.append(rule)

    def _get_is_match(self, node: XTreeNode) -> bool:
        for rule in self._rules:
            if rule.is_match(node, self._match_value):
                return True
        return False

    def _get_is_match_all(self, node: XTreeNode) -> bool:
        if not self._rules:
            return False
        for rule in self._rules:
            if not rule.is_match(node, self._match_value):
                return False

        return True

    def find_node(self, node: XTreeNode) -> XTreeNode | None:
        """
        Finds a node in the tree

        Args:
            node (XTreeNode): Node to find

        Returns:
            XTreeNode | None: The node if found, None otherwise
        """
        if self._match_all:
            found = self._get_is_match_all(node)
        else:
            found = self._get_is_match(node)
        if found:
            return node

        for i in range(node.getChildCount()):
            result = self.find_node(node.getChildAt(i))
            if result is not None:
                return result
        return None

    # endregion Methods

    # region Properties
    @property
    def match_all(self) -> bool:
        """Gets or sets whether all rules must match for a node to be considered a match."""
        return self._match_all

    @match_all.setter
    def match_all(self, value: bool) -> None:
        self._match_all = value

    @property
    def match_value(self) -> Any:
        """Gets or sets the value to match."""
        return self._match_value

    @match_value.setter
    def match_value(self, value: Any) -> None:
        self._match_value = value

    # endregion Properties
