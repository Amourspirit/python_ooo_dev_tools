from __future__ import annotations
from typing import Iterable, List, TYPE_CHECKING

import uno
from com.sun.star.beans import XMultiPropertySet

from ooodev.events.args.generic_args import GenericArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import gen_util as gUtil
from ooodev.loader import lo as mLo
from ooodev.adapter.beans.properties_change_events import PropertiesChangeEvents
from ooodev.adapter.beans.properties_change_listener import PropertiesChangeListener

if TYPE_CHECKING:
    from ooodev.utils.type_var import EventArgsCallbackT
    from ooodev.utils.type_var import ListenerEventCallbackT


class PropertiesChangeImplement:
    """
    Class for managing Properties Change Events.

    This class can be used to extend a class that already implements or uses ``com.sun.star.beans.XMultiPropertySet`` in some way.
    """

    def __init__(
        self,
        component: XMultiPropertySet,
        trigger_args: GenericArgs | None = None,
        cb: ListenerEventCallbackT | None = None,
        listener: PropertiesChangeListener | None = None,
    ) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Component that implements ``com.sun.star.beans.XMultiPropertySet``.
            trigger_args (Dict[str, Any], optional): Generic Arguments to pass to instances of ``PropertyChangeEvents``
                when they are created. Defaults to None.
        """
        self.__callback = cb
        self.__component = mLo.Lo.qi(XMultiPropertySet, component, True)
        if listener:
            self.__listener = listener
        else:
            self.__listener = PropertiesChangeListener(trigger_args=trigger_args)
        self.__current_names = set()
        self.__name = gUtil.Util.generate_random_string(10)
        self.__events = PropertiesChangeEvents(trigger_args=trigger_args, listener=self.__listener)

    def add_event_properties_change(self, names: Iterable[str], cb: EventArgsCallbackT) -> None:
        """
        Add properties to listen for changes.

        Args:
            names (Iterable[str]): One or more property names to listen for changes.
            cb (EventArgsCallbackT): Callback that is invoked when an event is triggered.

        Raises:
            ValueError: If names is empty.

        Returns:
            None:

        Note:
            The callback ``EventArgs.event_data`` will contain a tuple of ``com.sun.star.beans.PropertyChangeEvent`` objects.

            Each time this method is called, the previous names are removed and the new names are added.
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="add_event_properties_change")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        if not names:
            raise ValueError("names cannot be empty")

        self.remove_event_properties_listener()

        # ensure no duplicates
        self.__current_names.update(names)

        uno_strings = uno.Any("[]string", tuple(self.__current_names))  # type: ignore
        # self.__component.addPropertiesChangeListener(tuple(self.__current_names), self.__listener)
        uno.invoke(self.__component, "addPropertiesChangeListener", (uno_strings, self.__listener))
        self.__events.add_event_properties_change(cb=cb)

    def fire_event_properties_change(self, names: Iterable[str]) -> None:
        """
        Fires a sequence of PropertyChangeEvents

        Args:
            names (Iterable[str]): Sequence of property names to fire event for.

        Returns:
            None:
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="fire_event_properties_change")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None
        if not self.__current_names:
            # no names means no listener
            return

        if not names:
            raise ValueError("names cannot be empty")

        def get_existing_names(lst: Iterable[str]) -> List[str]:
            result = []
            for name in lst:
                if name in self.__current_names:
                    result.append(name)
            return result

        existing_names = get_existing_names(names)
        if not existing_names:
            return

        uno_strings = uno.Any("[]string", tuple(existing_names))  # type: ignore
        # self.__component.firePropertiesChangeEvent(tuple(existing_names), self.__listener)
        uno.invoke(self.__component, "firePropertiesChangeEvent", (uno_strings, self.__listener))

    def remove_event_properties_listener(self) -> None:
        """
        Remove Properties Listener

        Returns:
            None:
        """
        if self.__callback:
            args = ListenerEventArgs(source=self.__name, trigger_name="remove_event_properties_listener")
            self.__callback(self, args)
            if args.remove_callback:
                self.__callback = None

        if not self.__current_names:
            # no names means no listener
            return

        self.__component.removePropertiesChangeListener(self.__listener)
        self.__current_names.clear()
        self.__events.events_listener_properties_change.clear()

    @property
    def events_listener_properties_change_implement(self) -> PropertiesChangeListener:
        """
        Returns listener
        """
        return self.__listener
