from __future__ import annotations
from typing import Any


class DocCommonPartial:
    """Partial Class used for common document operations."""

    def __init__(self, component: Any) -> None:
        """
        Constructor.

        Args:
            component (Any): Any Uno Component that supports ``XStorageBasedLibraryContainer`` interface.
        """
        self.__component = component

    def __bool__(self) -> bool:
        return True

    @property
    def runtime_uid(self) -> str:
        """
        Gets the runtime id such as 1

        Returns:
            str: The runtime id.
        """
        return self.__component.RuntimeUID

    @property
    def string_value(self) -> str:
        """
        Gets the string value of the doc such as ``'vnd.sun.star.tdoc:/1/'``

        Returns:
            str: The string value.
        """
        return self.__component.StringValue
