from __future__ import annotations
from typing import Any, TYPE_CHECKING
import re
import contextlib

if TYPE_CHECKING:
    from com.sun.star.awt.tree import XTreeNode


class RuleTextRegex:
    """Rule for matching a node's text value with a regular expression pattern."""

    def __init__(self, regex: re.Pattern | str = "") -> None:
        """
        Constructor

        Arguments:
            regex (re.Pattern | str): Regular expression pattern to match.
                If empty string, then the ``match_value`` of ``is_match)()`` must be a regular expression pattern.
                If included then the ``match_value`` of ``is_match()`` is ignored. Defaults to ``""``.
        """
        if isinstance(regex, str):
            if regex:
                self._regex = re.compile(regex)
            else:
                self._regex = None
        else:
            self._regex = regex

    def is_match(self, node: XTreeNode, match_value: Any) -> bool:
        """
        Gets the node's text value and compares it with the match value.

        Args:
            node (XTreeNode): Tree node to check.
            match_value (Any): Value to match. Must be a regular expression pattern to match.

        Returns:
            bool: True if the node's text value matches the match value; Otherwise, False.
        """
        if self._regex:
            rx = self._regex
        else:
            try:
                if isinstance(match_value, str):
                    rx = re.compile(match_value)
                else:
                    rx = match_value
            except Exception:
                return False
            if not isinstance(rx, re.Pattern):
                return False
        with contextlib.suppress(Exception):
            node_text = node.getDisplayValue()
            return rx.search(node_text) is not None
        return False
