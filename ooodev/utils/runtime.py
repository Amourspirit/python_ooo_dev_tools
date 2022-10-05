from __future__ import annotations
from typing import Any
from . import lo as mLo

from com.sun.star.uno import XInterface

class Runtime:
    """Represents some of the UnoRuntime methods of Java"""
    @staticmethod
    def are_same(obj1: Any, obj2: Any) -> bool:
        """
        Determines if two Uno object are the same.

        Args:
            obj1 (Any): First Uno object
            obj2 (Any): Second Uno Object

        Returns:
            bool: True if objects are the same; Otherwise False.
        """
        # https://wiki.openoffice.org/wiki/Documentation/DevGuide/WritingUNO/XInterface
        # In C++, two objects are the same if their XInterface are the same. The queryInterface() for XInterface will have to
        # be called on both. In Java, check for the identity by calling the runtime function
        # com.sun.star.uni.UnoRuntime.areSame().
        x1 = mLo.Lo.qi(XInterface, obj1)
        if x1 is None:
            return False
        x2 = mLo.Lo.qi(XInterface, obj2)
        if x2 is None:
            return False
        return id(x1) == id(x2)