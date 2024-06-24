from __future__ import annotations
from typing import Any, TYPE_CHECKING

if TYPE_CHECKING:
    from typing_extensions import Protocol
    from ooodev.io.log.named_logger import NamedLogger
else:
    Protocol = object
    NamedLogger = Any


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

    @property
    def log(self) -> NamedLogger:
        """
        Gets the logger.

        Returns:
            NamedLogger: The logger.
        """
        ...
