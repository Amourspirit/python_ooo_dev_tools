from __future__ import annotations
from typing import Any, overload, Iterable, TYPE_CHECKING
import uno

if TYPE_CHECKING:
    from com.sun.star.frame import XFrame
    from com.sun.star.beans import PropertyValue  # struct

    from typing_extensions import Protocol
else:
    Protocol = object


class DispatchPartialT(Protocol):
    # region dispatch_cmd()

    @overload
    def dispatch_cmd(self, cmd: str) -> Any:
        """
        Dispatches a LibreOffice command.

        Args:
            cmd (str): Command to dispatch such as ``GoToCell``. Note: cmd does not contain ``.uno:`` prefix.

        Returns:
            Any: A possible result of the executed internal dispatch. The information behind this any depends on the dispatch!
        """
        ...

    @overload
    def dispatch_cmd(self, cmd: str, props: Iterable[PropertyValue]) -> Any:
        """
        Dispatches a LibreOffice command.

        Args:
            cmd (str): Command to dispatch such as ``GoToCell``. Note: cmd does not contain ``.uno:`` prefix.
            props (PropertyValue, optional): properties for dispatch.

        Returns:
            Any: A possible result of the executed internal dispatch. The information behind this any depends on the dispatch!
        """
        ...

    @overload
    def dispatch_cmd(self, cmd: str, props: Iterable[PropertyValue], frame: XFrame) -> Any:
        """
        Dispatches a LibreOffice command.

        Args:
            cmd (str): Command to dispatch such as ``GoToCell``. Note: cmd does not contain ``.uno:`` prefix.
            props (PropertyValue, optional): properties for dispatch.
            frame (XFrame, optional): Frame to dispatch to.

        Returns:
            Any: A possible result of the executed internal dispatch. The information behind this any depends on the dispatch!
        """
        ...

    @overload
    def dispatch_cmd(self, cmd: str, *, frame: XFrame) -> Any:
        """
        Dispatches a LibreOffice command.

        Args:
            cmd (str): Command to dispatch such as ``GoToCell``. Note: cmd does not contain ``.uno:`` prefix.
            props (PropertyValue, optional): properties for dispatch.
            frame (XFrame, optional): Frame to dispatch to.

        Returns:
            Any: A possible result of the executed internal dispatch. The information behind this any depends on the dispatch!
        """
        ...

    # endregion dispatch_cmd()
