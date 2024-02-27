from __future__ import annotations
from typing import Dict, TypedDict
import uno
from com.sun.star.beans import XPropertySet

from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.beans.vetoable_change_events import VetoableChangeEvents


class _PropertyDict(TypedDict):
    property_name: str
    ref_count: int
    event: VetoableChangeEvents


class VetoableChangeCollection:
    """
    Class for managing Adding and removing Vetoable Change Listeners.
    """

    def __init__(self, component: XPropertySet, generic_args: GenericArgs | None = None) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Component that implements ``com.sun.star.beans.XPropertySet``.
            generic_args (Dict[str, Any], optional): Generic Arguments to pass to instances of ``PropertyChangeEvents``
                when they are created. Defaults to None.
        """
        self.__events: Dict[str, _PropertyDict] = {}
        self.__component = component
        self.__generic_args = generic_args

    def create_vetoable_change_events(self, name: str, **kwargs) -> VetoableChangeEvents:
        """
        Creates a vetoable change events that can be used to add/remove listeners.

        Any generic args passed into the constructor are also passed into the constructor of the ``VetoableChangeEvents``.

        Args:
            name (str): Property Name
            kwargs (Any): Keyword Arguments to pass to the ``GenericArgs`` constructor.

        Raises:
            ValueError: If name is empty

        Returns:
            VetoableChangeEvents: Vetoable Change Events
        """
        if self.__generic_args is None:
            extra = {}
        else:
            extra = self.__generic_args.kwargs.copy()
        extra.update(kwargs)
        if not name:
            raise ValueError("name cannot be empty")
        args = GenericArgs(component=self.__component, property_name=name, **extra)
        return VetoableChangeEvents(trigger_args=args)

    def add_vetoable_listener(self, name: str, events: VetoableChangeEvents | None = None) -> VetoableChangeEvents:
        """
        Add vetoable property listener

        Args:
            name (str): Property Name
            events (VetoableChangeEvents | None, optional): Instance to pass to UNO. Defaults to None.

        Raises:
            ValueError: If name is empty.
            ValueError: If name has already been added for the current listener.

        Returns:
            VetoableChangeEvents: Vetoable Change Events

        Note:
            The returned instance, in most cases, would be used to subscribe to callbacks.
        """
        if not name:
            raise ValueError("name cannot be empty")
        if not events:
            events = self.create_vetoable_change_events(name)

        # get the name from mangling
        events_name: str = events._VetoableChangeEvents__name  # type: ignore
        if events_name in self.__events:
            # increment the reference count
            current = self.__events[events_name]
            if current["property_name"] == name:
                raise ValueError(f'The property "{name}" is already being listened to.')
            current_events = current["event"]
            self.__component.addVetoableChangeListener(name, current_events.events_listener_vetoable_change)
            current["ref_count"] += 1
        else:
            # add the listener
            self.__component.addVetoableChangeListener(name, events.events_listener_vetoable_change)
            self.__events[events_name] = {"property_name": name, "ref_count": 1, "event": events}
        return self.__events[events_name]["event"]

    def remove_vetoable_listener(self, name: str, events: VetoableChangeEvents) -> bool:
        """
        Remove Property Listener

        Args:
            name (str): Property Name
            events (VetoableChangeEvents): Property Change Events

        Raises:
            ValueError: If name is empty

        Returns:
            bool: ``True`` on success; Otherwise, ``False``.
        """
        if not name:
            raise ValueError("name cannot be empty")
        # get the name from mangling
        events_name: str = events._VetoableChangeEvents__name  # type: ignore
        if events_name not in self.__events:
            return False
        current = self.__events[events_name]
        ref_count = current["ref_count"]
        result = False
        if ref_count > 1:
            current["ref_count"] -= 1
            self.__component.removeVetoableChangeListener(name, current["event"].events_listener_vetoable_change)
            result = True
        else:
            self.__component.removeVetoableChangeListener(name, current["event"].events_listener_vetoable_change)
            _ = self.__events.pop(events_name)
            result = True
        return result
