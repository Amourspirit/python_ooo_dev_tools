from __future__ import annotations
from typing import Any, overload, Iterable, TYPE_CHECKING
import uno
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.events.lo_events import observe_events

if TYPE_CHECKING:
    from com.sun.star.frame import XFrame
    from ooodev.events.events_t import EventsT
    from com.sun.star.beans import PropertyValue  # struct


class DispatchPartial:
    def __init__(self, lo_inst: LoInst, events: EventsT | None = None):
        self.__lo_inst = lo_inst  # may be used in future
        if events is None:
            self.__events = None
        else:
            self.__events = events.event_observer

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

    def dispatch_cmd(self, cmd: str, props: Iterable[PropertyValue] | None = None, frame: XFrame | None = None) -> Any:
        """
        Dispatches a LibreOffice command.

        Args:
            cmd (str): Command to dispatch such as ``GoToCell``. Note: cmd does not contain ``.uno:`` prefix.
            props (PropertyValue, optional): properties for dispatch.
            frame (XFrame, optional): Frame to dispatch to.

        Raises:
            CancelEventError: If Dispatching is canceled via event.
            DispatchError: If any other error occurs.

        Returns:
            Any: A possible result of the executed internal dispatch. The information behind this any depends on the dispatch!

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.lo_named_event.LoNamedEvent.DISPATCHING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.lo_named_event.LoNamedEvent.DISPATCHED` :eventref:`src-docs-event`

        Note:
            There are many dispatch command constants that can be found in :ref:`utils_dispatch` Namespace

            | ``DISPATCHING`` Event args data contains any properties passed in via ``props``.
            | ``DISPATCHED`` Event args data contains any results from the dispatch commands.

        See Also:
            - :ref:`ch04_dispatching`
            - `LibreOffice Dispatch Commands <https://wiki.documentfoundation.org/Development/DispatchCommands>`_
        """
        if self.__events is not None:
            with observe_events(self.__events):
                return self.__lo_inst.dispatch_cmd(cmd, props, frame)  # type: ignore
        else:
            return self.__lo_inst.dispatch_cmd(cmd, props, frame)  # type: ignore

    # endregion dispatch_cmd()
