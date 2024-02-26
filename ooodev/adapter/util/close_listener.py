from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from com.sun.star.util import XCloseListener
from ooodev.events.args.key_val_args import KeyValArgs

from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.adapter_base import AdapterBase

if TYPE_CHECKING:
    from com.sun.star.util import XCloseBroadcaster
    from com.sun.star.lang import EventObject


class CloseListener(AdapterBase, XCloseListener):
    """
    Makes it possible to receive events when an object is called for closing

    Such close events are broadcasted by a XCloseBroadcaster if somewhere tries to close it by calling XCloseable.close(). Listener can:

    If an event com.sun.star.lang.XEventListener.disposing() occurred, nobody called XCloseable.close() on listened object before. Then it's not allowed to break this request - it must be accepted!

    See Also:
        `API XCloseListener <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1util_1_1XCloseListener.html>`_
    """

    def __init__(self, trigger_args: GenericArgs | None = None, subscriber: XCloseBroadcaster | None = None) -> None:
        """
        Constructor

        Args:
            trigger_args (GenericArgs, Optional): Args that are passed to events when they are triggered.
            subscriber (XCloseBroadcaster, optional): An UNO object that implements the ``XCloseBroadcaster`` interface.
                If passed in then this listener instance is automatically added to it.
        """
        super().__init__(trigger_args=trigger_args)
        # doc parameter was renamed to subscriber in version 0.13.6
        if subscriber:
            subscriber.addCloseListener(self)

    def notifyClosing(self, event: EventObject) -> None:
        """
        Is invoked when the listened object is closed really.

        Now the listened object is closed really.
        Listener has to accept that; should deregister itself and release all references to it.
        It's not allowed nor possible to disagree with that by throwing any exception.

        If the event ``com.sun.star.lang.XEventListener.disposing()`` occurred before it must be accepted too.
        There exist no chance for a disagreement any more.
        """
        self._trigger_event("notifyClosing", event)

    def queryClosing(self, event: EventObject, gets_ownership: bool) -> None:
        """
        Is invoked when somewhere tries to close listened object

        Is called before ``XCloseListener.notifyClosing()``.
        Listener has the chance to break that by throwing a CloseVetoException.
        This exception must be passed to the original caller of ``XCloseable.close()`` without any interaction.

        The parameter ``gets_ownership`` regulate who has to try to close the listened object again,
        if this listener disagree with the request by throwing the exception.
        If it's set to ``False`` the original caller of ``XCloseable.close()`` will be the owner in every case.
        It's not allowed to call ``close()`` from this listener then.
        If it's set to ``True`` this listener will be the new owner if he throw the exception, otherwise not!
        If his still running processes will be finished he must call ``close()`` on listened object again then.

        If this listener doesn't disagree with th close request it depends from his internal implementation if he deregister itself at the listened object.
        But normally this must be done in ``XCloseListener.notifyClosing()``.

        Args:
            event (EventObject): Describes the source of the event (must be the listened object)
            gets_ownership (bool): ``True`` pass the ownership to this listener, if he throw the veto exception (otherwise this parameter must be ignored!)
                ``False`` forbids to grab the ownership for the listened close object and call close() on that any time.

        Raises:
            CloseVetoException: ``CloseVetoException``

        Note:
            This implementation raises a ``queryClosing`` event in the event system.
            The event data is a ``KeyValArgs`` instance with the following properties:
            ``key=gets_ownership``, ``value=gets_ownership`` and ``event_data=event``.
        """
        kv_args = KeyValArgs(self, key="gets_ownership", value=gets_ownership)
        kv_args.event_data = event
        self._trigger_event("queryClosing", kv_args)

    def disposing(self, event: EventObject) -> None:
        """
        Gets invoked when the broadcaster is about to be disposed.

        All listeners and all other objects, which reference the broadcaster
        should release the reference to the source. No method should be invoked
        anymore on this object ( including ``XComponent.removeEventListener()`` ).

        This method is called for every listener registration of derived listener
        interfaced, not only for registrations at ``XComponent``.
        """
        # from com.sun.star.lang.XEventListener
        self._trigger_event("disposing", event)
