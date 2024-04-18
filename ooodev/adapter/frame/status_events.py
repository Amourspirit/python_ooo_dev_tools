from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING, Set
import uno
from com.sun.star.frame import XDispatch

from ooodev.events.args.generic_args import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.adapter.frame.status_listener import StatusListener

if TYPE_CHECKING:
    from com.sun.star.util import URL
    from ooodev.utils.type_var import EventArgsCallbackT


class StatusEvents:
    """
    Class for managing Status Events.
    """

    def __init__(self, subscriber: XDispatch, trigger_args: GenericArgs | None = None) -> None:
        """
        Constructor

        Args:
            subscriber (XDispatch): An UNO object that implements the ``XDispatch`` interface.
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
        """
        self._urls: Set[URL] = set()
        self.__subscriber = subscriber
        self.__listener = StatusListener(trigger_args=trigger_args)
        self.__name = gUtil.Util.generate_random_string(10)

    # region Manage Events
    def add_event_status_changed(self, url: URL, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the status of the feature changes.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.frame.FeatureStateEvent`` struct.
        """
        # FeatureStateEvent will contain a URL

        if url not in self._urls:
            self.__subscriber.addStatusListener(self.__listener, url)
            self._urls.add(url)
        self.__listener.on("statusChanged", cb)

    def add_event_status_events_disposing(self, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the broadcaster is about to be disposed.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.lang.EventObject`` struct.
        """

        self.__listener.on("disposing", cb)

    def remove_event_status_changed(self, url: URL, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        if url in self._urls:
            self.__subscriber.removeStatusListener(self.__listener, url)
            self._urls.remove(url)
        self.__listener.off("statusChanged", cb)

    def remove_event_border_resize_events_disposing(self, cb: EventArgsCallbackT) -> None:
        """
        Removes a listener for an event
        """
        self.__listener.off("disposing", cb)

    # endregion Manage Events


def on_lazy_cb(source: Any, event: ListenerEventArgs) -> None:
    """
    Callback that is invoked when an event is added or removed.

    This method is generally used to add the listener to the component in a lazy manner.
    This means this callback will only be called once in the lifetime of the component.

    Args:
        source (Any): Expected to be an instance of BorderResizeEvents that is a partial class of a component based class.
        event (ListenerEventArgs): Event arguments.

    Returns:
        None:

    Warning:
        This method is intended for internal use only.
    """
    # not used for StatusEvents
    pass
