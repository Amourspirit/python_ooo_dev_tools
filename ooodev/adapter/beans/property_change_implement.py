from __future__ import annotations
from typing import Dict, TYPE_CHECKING
import uno
from com.sun.star.beans import XPropertySet
from ooodev.events.args.generic_args import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.loader import lo as mLo
from ooodev.adapter.beans.property_change_collection import PropertyChangeCollection
from ooodev.adapter.beans.property_change_events import PropertyChangeEvents

if TYPE_CHECKING:
    from ooodev.utils.type_var import EventArgsCallbackT
    from ooodev.utils.type_var import ListenerEventCallbackT


class PropertyChangeImplement:
    """
    Class for managing Property Change Events.

    This class can be used to extend a class that already implements or uses ``com.sun.star.beans.XPropertySet`` in some way.
    """

    def __init__(
        self,
        component: XPropertySet,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
    ) -> None:
        """
        Constructor

        Args:
            component (XPropertySet): UNO Component that implements ``com.sun.star.beans.XPropertySet``.
            trigger_args (GenericArgs, optional): Args that are passed to events when they are triggered.
            cb (ListenerEventCallbackT | None, optional): Callback that is invoked when an event is added or removed.
        """
        self.__callback = cb
        self.__name = gUtil.Util.generate_random_string(10)
        self.__property_events: Dict[str, PropertyChangeEvents] = {}
        self.__dispose_events: Dict[str, PropertyChangeEvents] = {}
        ps = mLo.Lo.qi(XPropertySet, component, True)
        self.__property_collection = PropertyChangeCollection(component=ps, generic_args=trigger_args)

    # region Manage Events
    def add_event_property_change(self, name: str, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when property is changed.

        The callback ``EventArgs.event_data`` will contain a ``com.sun.star.beans.PropertyChangeEvent`` struct.

        Args:
            name (str): Property Name
            cb (EventArgsCallbackT): Callback

        Returns:
            None:
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="propertyChange")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        if name not in self.__property_events:
            event = self.__property_collection.add_property_listener(name=name)
            event.add_event_property_change(cb=cb)
            self.__property_events[name] = event

    def add_event_property_change_events_disposing(self, name: str, cb: EventArgsCallbackT) -> None:
        """
        Adds a listener for an event.

        Event is invoked when the property listener is about to be disposed.

        The callback ``EventArgs.event_data`` will contain a UNO ``com.sun.star.lang.EventObject`` struct.

        Args:
            name (str): Property Name
            cb (EventArgsCallbackT): Callback

        Returns:
            None:
        """

        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="disposing")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        if name not in self.__dispose_events:
            event = self.__property_collection.add_property_listener(name=name)
            event.add_event_property_change_events_disposing(cb=cb)
            self.__dispose_events[name] = event

    def remove_event_property_change(self, name: str) -> None:
        """
        Removes a listener for an event

        Args:
            name (str): Property Name

        Returns:
            None:
        """

        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="propertyChange", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        if name in self.__property_events:
            event = self.__property_events[name]
            self.__property_collection.remove_property_listener(name=name, events=event)
            self.__property_events.pop(name)

    def remove_event_property_change_events_disposing(self, name: str) -> None:
        """
        Removes a listener for an event

        Args:
            name (str): Property Name

        Returns:
            None:
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="disposing", is_add=False)
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        if name in self.__dispose_events:
            event = self.__dispose_events[name]
            self.__property_collection.remove_property_listener(name=name, events=event)
            self.__dispose_events.pop(name)

    # endregion Manage Events
