from __future__ import annotations
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from typing_extensions import Protocol
else:
    Protocol = object


class DocCommonPartialT(Protocol):

    @property
    def runtime_uid(self) -> str:
        """
        Gets the runtime id such as 1

        Returns:
            str: The runtime id.
        """
        ...

    @property
    def string_value(self) -> str:
        """
        Gets the string value of the doc such as ``'vnd.sun.star.tdoc:/1/'``

        Returns:
            str: The string value.
        """
        ...
