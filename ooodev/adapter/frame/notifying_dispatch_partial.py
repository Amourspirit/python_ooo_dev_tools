from __future__ import annotations
from enum import unique
from typing import Any, cast, TYPE_CHECKING, Set
import uno
from com.sun.star.frame import XNotifyingDispatch

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo


from ooodev.adapter.frame.dispatch_partial import DispatchPartial
from ooodev.events.args.generic_args import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.adapter.frame.dispatch_result_listener import DispatchResultListener

if TYPE_CHECKING:
    from com.sun.star.util import URL
    from com.sun.star.beans import PropertyValue
    from ooodev.utils.type_var import EventArgsCallbackT
    from ooodev.utils.type_var import UnoInterface


class NotifyingDispatchPartial(DispatchPartial):
    """
    Class for managing Status Events.
    """

    def __init__(self, component: XNotifyingDispatch, interface: UnoInterface | None = XNotifyingDispatch) -> None:
        """
        Constructor

        Args:
            subscriber (XNotifyingDispatch): An UNO object that implements the ``XNotifyingDispatch`` interface.
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
        """
        DispatchPartial.__init__(self, component, interface)
        self._NotifyingDispatchPartial__listeners = {}

        def _on_dispatched(*args, **kwargs):
            # don't hold on to the listener after the event is dispatched.
            # A new one will be created with the next dispatch.
            unique_id = kwargs.pop("unique_id", "")
            if unique_id:
                _ = self._NotifyingDispatchPartial__listeners.pop(unique_id, None)

        self.__fn_dispatched = _on_dispatched
        self.__component = component

    # region XNotifyingDispatch
    def dispatch_with_notification(self, *args: PropertyValue, url: URL, cb: EventArgsCallbackT) -> None:
        """
        Do the same like ``XDispatch.dispatch()`` but notifies listener in every case.

        Should be used if result must be known.

        The call back ``event_data`` will contain a UNO ``com.sun.star.frame.DispatchResultEvent`` struct.

        Args:
            url (URL): Specifies full parsed URL describes the feature which should be dispatched (executed).
            args (PropertyValue, optional): Optional arguments for this request.
            cb (EventArgsCallbackT): Callback that is invoked when an event is dispatched.

        Note:
            The Callback, ``cb``, is only a one-time event for each call to this method.
            If needed you can use the same callback for multiple calls to this method.
        """
        unique_id = gUtil.Util.generate_random_string(10)
        g_args = GenericArgs(unique_id=unique_id)
        listener = DispatchResultListener(trigger_args=g_args)
        listener.on("dispatchFinished", cb)
        listener.on("dispatchFinished", self.__fn_dispatched)
        self._NotifyingDispatchPartial__listeners[unique_id] = listener
        self.__component.dispatchWithNotification(url, args, listener)

    # endregion XNotifyingDispatch
